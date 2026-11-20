// Jordan Fein — Sheet → project pages sync
// Keeps every work/[slug]/index.html in the GitHub repo in step with the Credits column,
// and creates a project page for any new sheet row that has a video but no page yet.
//
// SETUP (one time)
// 1. Paste this file into the same Apps Script project as youtube-plays.gs (Extensions → Apps Script → + → Script).
// 2. Project Settings (gear) → Script Properties → Add property:
//      GITHUB_TOKEN = your fine-grained GitHub token (repo jordanfein1/jordanfein.com, "Contents: Read and write")
// 3. Select function "installSyncTrigger" and click Run once (approve the permissions).
//    After that it checks the sheet every 10 minutes and only touches GitHub when something changed.
//    To run it right away, reload the sheet and use the menu: Jordan Fein → Sync project pages now.

const REPO = 'jordanfein1/jordanfein.com';
const BRANCH = 'main';
const SKIP = ['the-blessing', 'the-blessing-trailer']; // custom layout, edited by hand
const MUX_TEMPLATE = { slug: 'rakuten', title: 'Rakuten', id: 'igPdGj25X2IGq6vMT7Cb6DlN1X2cDke7R001Hu9sK00Iw', cat: 'Commercial' };
const YT_TEMPLATE = { slug: 'kat-dennings', title: 'Kat Dennings', id: 'aLrbP-ujaI4', cat: 'Editorial' };

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Jordan Fein').addItem('Sync project pages now', 'syncPagesNow').addToUi();
}

function installSyncTrigger() {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'syncPages').forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('syncPages').timeBased().everyMinutes(10).create();
  syncPagesNow();
}

function syncPagesNow() { PropertiesService.getScriptProperties().deleteProperty('SYNC_HASH'); syncPages(); }

function syncPages() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return;
  try {
    const rows = readRows_();
    const hash = Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, JSON.stringify(rows)));
    const props = PropertiesService.getScriptProperties();
    if (props.getProperty('SYNC_HASH') === hash) return;

    const dirs = ghList_('work').filter(e => e.type === 'dir').map(e => e.name);
    const byId = {};
    const pages = {};
    dirs.forEach(slug => {
      const f = ghGet_('work/' + slug + '/index.html');
      if (!f) return;
      pages[slug] = f;
      const m = f.text.match(/data-mux="([^"]+)"/) || f.text.match(/youtube\.com\/embed\/([\w-]{11})/);
      if (m) byId[m[1]] = slug;
    });

    const log = [];
    const created = [];
    const used = {};
    const yearBySlug = {};
    rows.forEach(r => {
      if (!r.id) return;
      let s = slugify_(r.title); used[s] = (used[s] || 0) + 1; if (used[s] > 1) s += '-' + used[s]; if (s === 'the-blessing-trailer') s = 'the-blessing';
      let slug = byId[r.id];
      if (slug && r.year) yearBySlug[slug] = r.year;
      if (!slug) {
        if (dirs.indexOf(s) >= 0 || !r.title) return;
        const html = buildPage_(r, s, pages);
        if (!html) return;
        ghPut_('work/' + s + '/index.html', html, null, 'Add project page: ' + r.title);
        dirs.push(s); byId[r.id] = s; created.push({ slug: s, row: r });
        log.push('created ' + s);
        return;
      }
      if (SKIP.indexOf(slug) >= 0) return;
      const f = pages[slug];
      const next = applyKeywords_(applyYear_(applyDescription_(applyCredits_(f.text, r.credits), r), r.year), r);
      if (next !== f.text) { ghPut_('work/' + slug + '/index.html', next, f.sha, 'Update credits: ' + r.title); log.push('updated ' + slug); }
    });

    if (created.length) { addToPagesSet_(created.map(c => c.slug)); addToSitemap_(created); }
    try { updateSitemapYears_(yearBySlug); } catch (e) { console.warn('sitemap years: ' + e); }
    try { updateSnapshot_(); log.push('snapshot'); } catch (e) { console.warn('snapshot: ' + e); }
    props.setProperty('SYNC_HASH', hash);
    props.setProperty('SYNC_LAST', new Date().toISOString() + ' ' + (log.join(', ') || 'no changes'));
    console.log(log.join('\n') || 'no changes');
  } finally { lock.releaseLock(); }
}

// ---------- sheet ----------
function readRows_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheets().find(s => String(s.getRange(1, 1).getValue()).trim() === 'Project');
  const v = sh.getDataRange().getDisplayValues();
  const h = v[0].map(x => String(x).trim().toLowerCase());
  const col = n => h.indexOf(n);
  const cT = col('project'), cId = col('playback id'), cCat = col('category'), cCr = col('credits'), cTh = col('thumbnail time'), cDesc = col('description');
  return v.slice(1).map(r => {
    const raw = String(r[cId] || '').trim();
    const yt = raw.match(/(?:youtu\.be\/|v=|embed\/|shorts\/)([\w-]{11})/);
    return {
      title: fixTitle_(r[cT]),
      id: yt ? yt[1] : raw,
      yt: !!yt,
      cat: (String(r[cCat] || '').split(',').map(x => x.trim()).filter(Boolean)[0]) || 'Commercial',
      year: /^\d{4}$/.test(String(r[h.indexOf('year')] || '').trim()) ? String(r[h.indexOf('year')]).trim() : '',
      credits: fixCredits_(fixTitle_(r[cT]), r[cCr]),
      thumb: cTh >= 0 ? secs_(r[cTh]) : 0,
      hidden: /^(y|yes|true|x|1)$/i.test(String(r[h.indexOf('hide from all')] || '').trim()),
      desc: cDesc >= 0 ? String(r[cDesc] || '').replace(/\s+/g, ' ').trim() : '',
      keywords: h.indexOf('keywords') >= 0 ? String(r[h.indexOf('keywords')] || '').replace(/\s+/g, ' ').trim() : '',
    };
  }).filter(r => r.title);
}

function secs_(t) {
  t = String(t || '').trim(); if (!t) return 0;
  const p = t.split(':'); let s = 0;
  p.forEach(x => { s = s * 60 + (parseFloat(x) || 0); });
  return Math.round(s * 100) / 100;
}

function slugify_(s) {
  return String(s || '').toLowerCase().normalize('NFKD').replace(/[\u0300-\u036f]/g, '').replace(/&/g, ' and ').replace(/[^a-z0-9]+/g, '-').replace(/^-+|-+$/g, '') || 'project';
}

// Same corrections the site's applySheet() makes, so project pages, llms.txt and the SEO block never disagree with the site
// (and the sync can't revert them). Fix the sheet cells themselves and these become no-ops.
function fixTitle_(t) { return String(t || '').replace(/Who Care Trailer/g, 'Who Cares Trailer').trim(); }
function fixCredits_(title, c) {
  c = String(c || '').replace(/Richardo Acosta/g, 'Ricardo Acosta').replace(/Lance Edmonds/g, 'Lance Edmands').replace(/Travel\+Leisure/g, 'Travel + Leisure').replace(/Production co\b\.?/gi, 'Production Co.');
  if (/^new york city ballet$/i.test(title)) c = c.replace(/Client: (?:Entertainment Weekly|People Magazine)/, 'Client: People').replace(/Awards: Webby[^\n]*/, "Recognition: 2026 Webby Winner + People's Voice Winner — Health, Wellness & Fitness");
  if (/^i survived a face transplant$/i.test(title)) c = c.replace(/Recognition: (?:2026 )?New York Emmy (?:Awards )?(?:Nominee|Winner)[^\n]*/, 'Recognition: 2026 New York Emmy Winner — Human Interest, Long Form Content');
  return c;
}

// ---------- credits ----------
function parseCredits_(text) {
  const out = [];
  String(text || '').split(/\r?\n/).forEach(line => {
    const m = line.match(/^\s*([^:]{1,40}):\s*(.*)$/);
    if (m) out.push({ label: m[1].trim(), value: m[2].trim() });
    else if (line.trim() && out.length) out[out.length - 1].value += '\n' + line.trim();
  });
  return out.filter(c => c.value).map(c => {
    let l = c.label;
    if (/^awards?$/i.test(l)) l = 'Recognition';
    if (/^production co\.?$/i.test(l)) l = 'Production Co.';
    if (/^directed by$/i.test(l)) l = 'Director';
    if (/^dp$/i.test(l)) l = 'DP';
    return { label: l, value: c.value.replace(/\s*;\s*/g, '\n') };
  });
}

// Real year from the sheet's Year column → uploadDate in the page's VideoObject (pages are created with a 2026-01-01 placeholder).
function applyYear_(html, year) {
  if (!/^\d{4}$/.test(year || '')) return html;
  return html.replace(/("uploadDate":")\d{4}(-01-01T00:00:00-05:00")/, '$1' + year + '$2');
}

// Same for <video:publication_date> in sitemap.xml.
function updateSitemapYears_(yearBySlug) {
  const slugs = Object.keys(yearBySlug);
  if (!slugs.length) return;
  const f = ghGet_('sitemap.xml');
  if (!f) return;
  let next = f.text;
  slugs.forEach(s => {
    next = next.replace(new RegExp('(<loc>https://jordanfein\\.com/work/' + s + '/</loc>[\\s\\S]*?<video:publication_date>)\\d{4}(-01-01T00:00:00-05:00)'), (m, a, b) => a + yearBySlug[s] + b);
  });
  if (next !== f.text) ghPut_('sitemap.xml', next, f.sha, 'Sitemap: publication years from sheet');
}

function esc_(s) { return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;'); }

function applyCredits_(html, creditsText) {
  const m = html.match(/<dl style="margin:0;">([\s\S]*?)<\/dl>/);
  if (!m) return html;
  const row = (m[1].match(/<div style="display:grid[^>]*>[\s\S]*?<\/div>/) || [])[0];
  if (!row) return html;
  const cr = parseCredits_(creditsText);
  const dirIdx = cr.findIndex(c => /^directors?$/i.test(c.label));
  const dir = dirIdx >= 0 ? cr.splice(dirIdx, 1)[0] : { label: 'Director', value: 'Jordan Fein' };
  const mk = c => row.replace(/(<dt[^>]*>)[^<]*(<\/dt>)/, '$1' + esc_(c.label) + '$2').replace(/(<dd[^>]*>)[^<]*(<\/dd>)/, '$1' + esc_(c.value) + '$2');
  const dl = '<dl style="margin:0;">' + [dir].concat(cr).map(mk).join('') + '</dl>';
  return html.replace(m[0], dl);
}

// ---------- new pages ----------
function buildPage_(r, slug, pages) {
  const T = r.yt ? YT_TEMPLATE : MUX_TEMPLATE;
  const src = pages[T.slug];
  if (!src) return null;
  let h = src.text;
  const cat = r.cat.replace(/\b\w/g, c => c.toUpperCase());
  const lc = cat.toLowerCase(), art = /^[aeiou]/.test(lc) ? 'an' : 'a';
  const tlc = T.cat.toLowerCase(), tart = /^[aeiou]/.test(tlc) ? 'an' : 'a';
  // JSON-LD blocks get raw text, everything else gets HTML-escaped text
  const parts = h.split(/(<script type="application\/ld\+json">[\s\S]*?<\/script>)/);
  h = parts.map(p => {
    const json = p.indexOf('application/ld+json') >= 0;
    const title = json ? JSON.stringify(r.title).slice(1, -1) : esc_(r.title);
    return p.split(T.id).join(r.id)
      .split('/work/' + T.slug + '/').join('/work/' + slug + '/')
      .split(', ' + tart + ' ' + tlc + ' directed').join(', ' + art + ' ' + lc + ' directed')
      .split('| ' + T.cat + ' Directed').join('| ' + cat + ' Directed')
      .split('"genre":["' + T.cat + '"]').join('"genre":["' + cat + '"]')
      .split('>' + T.cat + '</p>').join('>' + esc_(cat) + '</p>')
      .split('<video:description>' + T.cat).join('<video:description>' + cat)
      .split(T.title).join(title);
  }).join('');
  if (!r.yt && r.thumb) h = h.replace(/thumbnail\.jpg\?width=1600&(amp;)?time=[\d.]+/g, (m0, a) => 'thumbnail.jpg?width=1600&' + (a || '') + 'time=' + r.thumb);
  if (!r.yt) { const d = muxDuration_(r.id); h = h.replace(/,"duration":"PT[^"]*"/g, d ? ',"duration":"' + d + '"' : ''); }
  h = applyYear_(h, r.year);
  if (!r.yt && !mp4Ok_(r.id)) { // no static rendition on this asset yet: HLS only (enable "static renditions" in Mux, then re-sync)
    h = h.replace(/<source src="https:\/\/stream\.mux\.com\/[^"]+\/highest\.mp4" type="video\/mp4">/g, '');
    h = h.replace(/"contentUrl":"https:\/\/stream\.mux\.com\/([^"/]+)\/highest\.mp4"/g, '"contentUrl":"https://stream.mux.com/$1.m3u8"');
  }
  h = h.replace(/<li><a href="\/work\/[^"]+\/">[^<]*<\/a><\/li>/g, (li) => li.indexOf('/work/' + slug + '/') >= 0 ? '' : li);
  if (h.indexOf('/work/' + T.slug + '/') < 0) h = h.replace(/(<ul[^>]*grid-template-columns:repeat\(auto-fill[^>]*>)/, '$1<li><a href="/work/' + T.slug + '/">' + esc_(T.title) + '</a></li>');
  return applyKeywords_(applyDescription_(applyCredits_(h, r.credits), r), r);
}

// "Description" column (2-3 sentences about the project) → visible paragraph + meta/OG/Twitter description + VideoObject description.
// Empty cell: page keeps its generic copy. Google needs unique on-page text to index a page, not just credits.
function applyDescription_(html, r) {
  const d = r.desc;
  if (!d) return html;
  const metaD = esc_(d).replace(/"/g, '&quot;');
  const tail = ' Film and commercial director based in New York and Los Angeles.';
  const m1 = metaD.length > 140 ? metaD : metaD + tail;
  html = html.replace(/(<meta (?:name="description"|property="og:description"|name="twitter:description") content=")[^"]*(")/g, (m0, a, b) => a + m1 + b);
  html = html.replace(/(<script type="application\/ld\+json">\{[^\n]*?"@type":"VideoObject"[^\n]*?"description":")(?:[^"\\]|\\.)*(")/, (m0, a, b) => a + JSON.stringify(d).slice(1, -1) + b);
  const para = '<p data-jf-desc="1" style="margin:0;max-width:720px;font-size:19px;line-height:1.55;color:#f2efe9;">' + esc_(d) + '</p>';
  if (/<p data-jf-desc="1"[^>]*>[\s\S]*?<\/p>/.test(html)) return html.replace(/<p data-jf-desc="1"[^>]*>[\s\S]*?<\/p>/, para);
  return html.replace(/(\n(\s*)<section style="display:grid;grid-template-columns:repeat\(auto-fit)/, (m0, a, sp) => '\n' + sp + para + a);
}

// Optional "Keywords" column (comma-separated subjects: sport, brand, people, format) → VideoObject.keywords.
// Only metadata that matches what the page already shows; empty cell leaves the page alone.
function applyKeywords_(html, r) {
  const k = r.keywords; if (!k) return html;
  const val = JSON.stringify(k).slice(1, -1);
  if (/"@type":"VideoObject"[^\n]*?"keywords":"/.test(html)) return html.replace(/("@type":"VideoObject"[^\n]*?"keywords":")(?:[^"\\]|\\.)*(")/, (m0, a, b) => a + val + b);
  return html.replace(/("@type":"VideoObject"[^\n]*?"genre":\[[^\]]*\])/, (m0, a) => a + ',"keywords":"' + val + '"');
}

// Googlebot's Chrome can't play HLS, so watch pages also carry an MP4 <source> (Mux static rendition "highest.mp4").
function mp4Ok_(id) {
  try { return UrlFetchApp.fetch('https://stream.mux.com/' + id + '/highest.mp4', { method: 'head', muteHttpExceptions: true }).getResponseCode() === 200; } catch (e) { return false; }
}

function muxDuration_(id) {
  try {
    const m = UrlFetchApp.fetch('https://stream.mux.com/' + id + '.m3u8', { muteHttpExceptions: true }).getContentText();
    const uri = (m.match(/#EXT-X-STREAM-INF[^\n]*\n([^\n]+)/) || [])[1];
    if (!uri) return '';
    let s = 0;
    UrlFetchApp.fetch(uri.trim(), { muteHttpExceptions: true }).getContentText().replace(/#EXTINF:([\d.]+)/g, (a, b) => { s += parseFloat(b); return a; });
    s = Math.round(s); if (!s) return '';
    const mi = Math.floor(s / 60);
    return 'PT' + (mi ? mi + 'M' : '') + (s % 60) + 'S';
  } catch (e) { return ''; }
}

function addToPagesSet_(slugs) {
  const f = ghGet_('index.html');
  if (!f) return;
  const next = f.text.replace(/const PAGES = new Set\(\[([\s\S]*?)\]\)/, (m, inner) => {
    const have = (inner.match(/"([^"]+)"/g) || []).map(x => x.slice(1, -1));
    slugs.forEach(s => { if (have.indexOf(s) < 0) have.push(s); });
    return 'const PAGES = new Set([' + have.sort().map(s => '"' + s + '"').join(',') + '])';
  });
  if (next !== f.text) ghPut_('index.html', next, f.sha, 'Link new project pages: ' + slugs.join(', '));
}

function addToSitemap_(created) {
  const f = ghGet_('sitemap.xml');
  if (!f) return;
  const day = Utilities.formatDate(new Date(), 'America/New_York', 'yyyy-MM-dd');
  const x = s => esc_(s);
  const blocks = created.filter(c => f.text.indexOf('/work/' + c.slug + '/') < 0).map(c => {
    const r = c.row;
    const thumb = r.yt ? 'https://i.ytimg.com/vi/' + r.id + '/maxresdefault.jpg' : 'https://image.mux.com/' + r.id + '/thumbnail.jpg?width=1600&amp;time=' + (r.thumb || 1);
    const player = r.yt ? 'https://www.youtube.com/embed/' + r.id : 'https://player.mux.com/' + r.id;
    return '  <url>\n    <loc>https://jordanfein.com/work/' + c.slug + '/</loc>\n    <lastmod>' + day + '</lastmod>\n    <video:video>\n' +
      '      <video:thumbnail_loc>' + thumb + '</video:thumbnail_loc>\n      <video:title>' + x(r.title) + ' — Director: Jordan Fein</video:title>\n' +
      '      <video:description>' + x(r.cat) + ' directed by Jordan Fein.</video:description>\n' + (!r.yt && mp4Ok_(r.id) ? '      <video:content_loc>https://stream.mux.com/' + r.id + '/highest.mp4</video:content_loc>\n' : '') + '      <video:player_loc>' + player + '</video:player_loc>\n' +
      '      <video:publication_date>' + (r.year || '2026') + '-01-01T00:00:00-05:00</video:publication_date>\n    </video:video>\n  </url>\n';
  }).join('');
  if (!blocks) return;
  ghPut_('sitemap.xml', f.text.replace('</urlset>', blocks + '</urlset>'), f.sha, 'Sitemap: add new project pages');
}

// ---------- first-visit snapshot ----------
// Bakes a copy of the sheet into index.html so Featured Work shows instantly for first-time visitors,
// and points the hero image preload at the first Featured video.
function updateSnapshot_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheets().find(x => String(x.getRange(1, 1).getValue()).trim() === 'Project');
  const v = sh.getDataRange().getDisplayValues();
  while (v.length > 1 && v[v.length - 1].every(c => !String(c).trim())) v.pop();
  const csv = v.map(r => r.map(c => '"' + String(c).replace(/"/g, '""') + '"').join(',')).join('\n');
  const f = ghGet_('index.html');
  if (!f) return;
  let next = f.text.replace(/const SHEET_SNAPSHOT = "(?:[^"\\]|\\.)*";/, () => 'const SHEET_SNAPSHOT = ' + JSON.stringify(csv) + ';');
  const rows = readRows_();
  const hero = rows.find(r => r.id && !r.yt && !r.hidden);
  if (hero) {
    let t = hero.thumb || 0;
    if (/blessing/i.test(hero.title)) t = 12.6; else if (/^celebrity$/i.test(hero.title)) t = 3.2; else if (!t && /sofia vergara/i.test(hero.title)) t = 2;
    next = next.replace(/(data-jf-hero-hls="1" href="https:\/\/stream\.mux\.com\/)[A-Za-z0-9]+/g, (m, p) => p + hero.id);
    next = next.replace(/(data-jf-hero="(\d+)" href="https:\/\/image\.mux\.com\/)[^/]+(\/thumbnail\.webp\?width=)\d+&amp;time=[\d.]+/g, (m, a, w, b) => a + hero.id + b + w + '&amp;time=' + t);
  }
  const seoRows = v.slice(1).map(r => ({ title: r[v[0].indexOf('Project')], id: String(r[v[0].indexOf('Playback ID')] || '').trim(), cats: r[v[0].indexOf('Category')], credits: r[v[0].indexOf('Credits')] }));
  const ldRows = v.slice(1).map(r => ({ title: r[v[0].indexOf('Project')], id: r[v[0].indexOf('Playback ID')], cats: r[v[0].indexOf('Category')], year: r[v[0].indexOf('Year')], thumb: r[v[0].indexOf('Thumbnail time')] }));
  next = next.replace(/(<script type="application\/ld\+json" id="jf-works">)[\s\S]*?(<\/script>)/, (m, a, b) => a + buildWorksLd_(ldRows) + b);
  next = next.replace(/(<div data-jf-seo="1"[^>]*>)[\s\S]*?(<span data-jf-seo-end="1"><\/span><\/div>)/, (m, a, b) => a + buildSeoBlock_(seoRows) + b);
  if (next !== f.text) ghPut_('index.html', next, f.sha, 'Refresh first-visit sheet snapshot');
  const lf = ghGet_('llms.txt');
  const llms = buildLlms_(ldRows.map((r, i) => Object.assign({}, r, { credits: seoRows[i].credits })));
  if (!lf || lf.text !== llms) ghPut_('llms.txt', llms, lf ? lf.sha : null, 'Refresh llms.txt');
}

function buildSeoBlock_(rows) {
  const esc = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  const fix = (s) => String(s || '').replace(/Richardo Acosta/g, 'Ricardo Acosta').replace(/Lance Edmonds/g, 'Lance Edmands').replace(/Who Care Trailer/g, 'Who Cares Trailer').replace(/Travel\+Leisure/g, 'Travel + Leisure').replace(/Production co\b\.?/gi, 'Production Co.');
  const slugify = (s) => String(s || '').toLowerCase().normalize('NFKD').replace(/[\u0300-\u036f]/g, '').replace(/&/g, ' and ').replace(/[^a-z0-9]+/g, '-').replace(/^-+|-+$/g, '') || 'project';
  const used = {};
  const items = rows.map((r) => {
    const title = fix(r.title).trim();
    if (!title || !r.id) return '';
    let s = slugify(title); used[s] = (used[s] || 0) + 1; if (used[s] > 1) s += '-' + used[s]; if (s === 'the-blessing-trailer') s = 'the-blessing';
    const cats = String(r.cats || '').split(',').map((x) => x.trim()).filter(Boolean).map((x) => x.replace(/\b\w/g, (c) => c.toUpperCase())).filter((x, i, a) => a.indexOf(x) === i);
    const lines = fixCredits_(title, fix(r.credits)).split(/\r?\n/);
    const parts = []; const awards = []; let inAwards = false;
    lines.forEach((ln) => {
      const m = ln.match(/^\s*([^:]{1,40}):\s*(.*)$/);
      if (m) {
        const label = m[1].trim(), val = m[2].trim();
        inAwards = /^awards?$/i.test(label);
        if (inAwards) { if (val) awards.push(val); return; }
        if (val && /^(client|production co\.?|agency|publications?|broadcast)$/i.test(label)) parts.push(label + ': ' + val);
      } else if (inAwards && ln.trim()) awards.push(ln.trim());
    });
    const bits = [cats.join(', ')].concat(parts).concat(awards.length ? ['Awards: ' + awards.join('; ')] : []).filter(Boolean);
    return '<li><a href="/work/' + s + '/">' + esc(title) + '</a>' + (bits.length ? ' — ' + esc(bits.join(' · ')) : '') + '</li>';
  }).filter(Boolean);
  return '<h1>Jordan Fein, Director — Commercial, Documentary, Sports &amp; Unscripted</h1><p>Jordan Fein is a freelance commercial and documentary director based in New York and Los Angeles. He directs commercials, documentaries, sports and unscripted series for brands, agencies and publishers. Selected work directed by Jordan Fein:</p><ul>' + items.join('') + '</ul>';
}

// Home page work list: links to the /work/ watch pages only. No VideoObject here — the home page is not a watch page,
// and VideoObjects on it trigger Search Console's "Video isn't on a watch page".
function buildWorksLd_(rows) {
  const fix = (s) => String(s || '').replace(/Who Care Trailer/g, 'Who Cares Trailer');
  const slugify = (s) => String(s || '').toLowerCase().normalize('NFKD').replace(/[\u0300-\u036f]/g, '').replace(/&/g, ' and ').replace(/[^a-z0-9]+/g, '-').replace(/^-+|-+$/g, '') || 'project';
  const used = {}; const items = [];
  rows.forEach((r) => {
    const title = fix(r.title).trim(); const raw = String(r.id || '').trim();
    if (!title || !raw) return;
    let s = slugify(title); used[s] = (used[s] || 0) + 1; if (used[s] > 1) s += '-' + used[s]; if (s === 'the-blessing-trailer') s = 'the-blessing';
    items.push({ '@type': 'ListItem', position: items.length + 1, name: title, url: 'https://jordanfein.com/work/' + s + '/' });
  });
  return JSON.stringify({ '@context': 'https://schema.org', '@type': 'ItemList', name: "Jordan Fein — Director's work", itemListOrder: 'https://schema.org/ItemListOrderAscending', numberOfItems: items.length, itemListElement: items }).replace(/</g, '\\u003c');
}

function buildLlms_(rows) {
  const fix = (s) => String(s || '').replace(/Richardo Acosta/g, 'Ricardo Acosta').replace(/Lance Edmonds/g, 'Lance Edmands').replace(/Who Care Trailer/g, 'Who Cares Trailer').replace(/Travel\+Leisure/g, 'Travel + Leisure').replace(/Production co\b\.?/gi, 'Production Co.');
  const slugify = (s) => String(s || '').toLowerCase().normalize('NFKD').replace(/[\u0300-\u036f]/g, '').replace(/&/g, ' and ').replace(/[^a-z0-9]+/g, '-').replace(/^-+|-+$/g, '') || 'project';
  const used = {}; const lines = [];
  rows.forEach((r) => {
    const title = fix(r.title).trim(); if (!title || !String(r.id || '').trim()) return;
    let s = slugify(title); used[s] = (used[s] || 0) + 1; if (used[s] > 1) s += '-' + used[s]; if (s === 'the-blessing-trailer') s = 'the-blessing';
    const cats = String(r.cats || '').split(',').map((x) => x.trim()).filter(Boolean).map((x) => x.replace(/\b\w/g, (c) => c.toUpperCase())).filter((x, i, a) => a.indexOf(x) === i);
    let director = 'Jordan Fein'; const parts = []; const awards = []; let inAwards = false;
    fixCredits_(title, fix(r.credits)).split(/\r?\n/).forEach((ln) => {
      const m = ln.match(/^\s*([^:]{1,40}):\s*(.*)$/);
      if (m) {
        const label = m[1].trim(), val = m[2].trim();
        inAwards = /^awards?$/i.test(label);
        if (inAwards) { if (val) awards.push(val); return; }
        if (!val) return;
        if (/^(directors?|directed by)$/i.test(label)) director = val.replace(/\s*;\s*/g, ', ');
        else parts.push(label + ': ' + val);
      } else if (inAwards && ln.trim()) awards.push(ln.trim());
    });
    const bits = [cats.join(' / '), 'Director: ' + director].concat(parts);
    if (awards.length) bits.push('Awards: ' + awards.join('; '));
    lines.push('- [' + title + '](https://jordanfein.com/work/' + s + '/): ' + bits.filter(Boolean).join('. ') + '.');
  });
  return ['# Jordan Fein — Director', '',
    '> Jordan Fein is a freelance commercial and documentary director based in New York and Los Angeles. He directs commercials, documentaries, sports and unscripted series for brands, agencies and publishers. Texas-raised, NYU Tisch alumnus, New York Emmy and Webby winner.', '',
    '## What he directs',
    '- Commercials and branded content for brands and agencies, including U.S. Soccer, Volvo, Land Rover, Chevy, Verizon, Squarespace, Adobe, Starbucks, Penn State Health and Formula 1',
    '- Documentary: The Blessing (feature documentary; PBS, WORLD, NHK Japan) and short documentaries for People, including the unscripted series I Survived',
    '- Celebrity and editorial film for People, Entertainment Weekly, InStyle and Travel + Leisure',
    '- Sports: Team USA Soccer for U.S. Soccer, Formula 1 and Formula 1 drivers, Penn State Health hockey',
    '- Health and real-people storytelling for hospital systems and health brands', '',
    '## Recognition',
    '- New York Emmy Award, Human Interest – Long Form Content (2026), for I Survived a Face Transplant (People)',
    "- Webby Award and People's Voice Award, Health, Wellness & Fitness (2026), for the New York City Ballet film (People)",
    '- Clio, Cannes Lions and One Show for Volvo "Highway Robbery" (co-directed)',
    '- Shorty Impact Award for Starbucks x SELF',
    '- The Blessing: Grand Jury Prize at the Dallas International Film Festival and five further festival awards',
    '- Named a Filmmaker to Watch by Independent Magazine; member of the Directors Guild of America', '',
    '## Contact and profiles',
    '- Website: https://jordanfein.com/',
    '- Email: jdfein@gmail.com',
    '- IMDb: https://www.imdb.com/name/nm3583341/',
    '- Vimeo: https://vimeo.com/jordanfein',
    '- Instagram: https://www.instagram.com/jordanfein/',
    '- LinkedIn: https://www.linkedin.com/in/jordan-fein-10aa011b/',
    '- Wikidata: https://www.wikidata.org/wiki/Q141601492', '',
    '## Key pages',
    '- Library of all work: https://jordanfein.com/library/',
    '- Commercial: https://jordanfein.com/library/commercial/',
    '- Real people and documentary: https://jordanfein.com/library/real-people/',
    '- Celebrity: https://jordanfein.com/library/celebrity/',
    '- Editorial: https://jordanfein.com/library/editorial/',
    '- Health: https://jordanfein.com/library/health/',
    '- Automotive: https://jordanfein.com/library/automotive/', '',
    '## Work (' + lines.length + ' projects, each with its own page)', ''].concat(lines).join('\n') + '\n';
}

// ---------- GitHub ----------
function gh_(method, path, body) {
  const token = PropertiesService.getScriptProperties().getProperty('GITHUB_TOKEN');
  if (!token) throw new Error('Add GITHUB_TOKEN in Project Settings → Script Properties');
  let res, code;
  for (let n = 0; ; n++) {
    res = UrlFetchApp.fetch('https://api.github.com/repos/' + REPO + '/contents/' + encodeURI(path) + (method === 'get' ? '?ref=' + BRANCH : ''), {
      method: method, muteHttpExceptions: true, contentType: 'application/json',
      headers: { Authorization: 'Bearer ' + token, Accept: 'application/vnd.github+json', 'X-GitHub-Api-Version': '2022-11-28' },
      payload: body ? JSON.stringify(body) : undefined,
    });
    code = res.getResponseCode();
    if ((code < 500 && code !== 429) || n >= 2) break; // retry GitHub outages / rate limits twice
    Utilities.sleep(3000 * (n + 1));
  }
  if (code === 404 && method === 'get') return null;
  if (code >= 300) throw new Error('GitHub ' + code + ' on ' + path + ': ' + res.getContentText().slice(0, 300));
  return JSON.parse(res.getContentText());
}

function ghList_(path) { return gh_('get', path) || []; }

function ghGet_(path) {
  const j = gh_('get', path);
  if (!j || !j.content) return null;
  const text = Utilities.newBlob(Utilities.base64Decode(j.content.replace(/\n/g, ''))).getDataAsString('UTF-8');
  return { text: text, sha: j.sha };
}

function ghPut_(path, text, sha, message) {
  const body = { message: message, branch: BRANCH, content: Utilities.base64Encode(text, Utilities.Charset.UTF_8) };
  if (sha) body.sha = sha;
  gh_('put', path, body);
}
