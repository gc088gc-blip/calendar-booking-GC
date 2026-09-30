/**
 * 行事曆與預約系統 — Google Apps Script 後端
 * 開源專案：一個人用的行程管理＋對外預約網頁
 * 資料：行程存在你的 Google 日曆；設定與收件匣存在 Google 試算表「預約系統」
 */

var TZ = 'Asia/Taipei';
var RANGE_END = '2027-12-31';
var STEP = 30;                       // 時段間隔（分鐘）
var INBOX = '收件匣';
var CFG_SHEET = '_設定';
var COLS = ['ID', '送出時間', '狀態', '類型', '對象', '名字', '聯絡方式', '想約時間', '多久', '要幹嘛', '約哪', '地點', '會議連結', '行事曆ID', '事件ID'];
var WDN = ['日', '一', '二', '三', '四', '五', '六'];
var BYDAY = ['SU', 'MO', 'TU', 'WE', 'TH', 'FR', 'SA'];
var AUD_NAME = { friend: '朋友', public: '新朋友' };

/** 第一次執行 setup 時會建立的行事曆，之後都在後台「設定 → 行事曆」改 */
var DEFAULT_CALS = [
  { name: '私人', color: '#4f7fd6', keywords: ['朋友', '家人'] },
  { name: '私事·工作相關', color: '#9467c9', keywords: ['會議', '開會'] },
  { name: '副業', color: '#cc7433', keywords: ['諮詢', '面談'] },
  { name: '正職', color: '#3f9a5e', keywords: ['上班', '公司'] }
];

/* ======================= 初次設定（在編輯器手動執行一次） ======================= */

function setup() {
  var props = PropertiesService.getScriptProperties();
  var ss, ssid = props.getProperty('ssid');
  if (ssid) ss = SpreadsheetApp.openById(ssid);
  else { ss = SpreadsheetApp.create('預約系統'); props.setProperty('ssid', ss.getId()); }

  var inbox = ss.getSheetByName(INBOX);
  if (!inbox) { inbox = ss.getSheets()[0]; inbox.setName(INBOX); }
  if (inbox.getLastRow() === 0) {
    inbox.appendRow(COLS);
    inbox.setFrozenRows(1);
    inbox.getRange(1, 1, 1, COLS.length).setFontWeight('bold').setBackground('#1d1c1a').setFontColor('#ffffff');
    inbox.setColumnWidth(10, 280);
  }
  var cs = ss.getSheetByName(CFG_SHEET) || ss.insertSheet(CFG_SHEET);
  cs.hideSheet();

  if (!cs.getRange('A1').getValue()) {
    var cals = DEFAULT_CALS.map(function (c) {
      var cal = CalendarApp.getCalendarsByName(c.name)[0] || CalendarApp.createCalendar(c.name, { color: c.color, timeZone: TZ });
      return { id: cal.getId(), name: c.name, color: c.color, keywords: c.keywords };
    });
    saveConfig_(defaultConfig_(cals));
  }
  var cfg = getConfig_();
  Logger.log('完成！試算表：' + ss.getUrl());
  Logger.log('朋友連結密鑰：' + cfg.audiences.friend.token);
}

function defaultConfig_(cals) {
  return {
    calendars: cals,
    bookTo: { friend: cals[0].id, public: cals[2].id },
    busyAlso: [Session.getEffectiveUser().getEmail()],     // 主日曆的行程也算忙碌
    audiences: {
      friend: { minNoticeH: 4, token: newToken_() },
      public: { minNoticeH: 24 }
    },
    shareCals: [],     // 分享連結要給看的行事曆
    shareToken: newToken_(),
    open: {},          // 'YYYY-MM-DD': [{ s:'19:00', e:'22:00', loc:'both'|'online', aud:'all'|'friend'|'public' }]
    people: [],        // 用過的「對象」
    me: { name: '我', brand: '', tagline: '' },     // 顯示在預約頁上的名字與品牌
    notifyEmail: Session.getEffectiveUser().getEmail(),
    publicUrl: '',
    shortFriend: '',
    shortPublic: '',
    shortShare: '',
    review: { friend: false, public: true },    // 新朋友的預約要先經過你同意
    notes: {},         // 每個月給預約者看的話：'YYYY-MM': { f: '給朋友的', p: '給新朋友的' }
    familyCal: '', familyName: ''   // 家庭日曆：勾「讓家人看到」的行程會複製一份放進去
  };
}

/* ======================= 網頁入口 ======================= */

function doGet(e) {
  var p = (e && e.parameter) || {};
  if (p.admin !== undefined) {
    if (!isOwner_()) return msgPage_('這個頁面只有本人能開。');
    var data = adminBoot();
    return page_('Admin', { here: selfUrl_(null, e), me: data.cfg.me, data: data }, data.cfg.me.name + ' 行事曆', true);
  }
  var cfg = getConfig_();
  if (p.v === 'f') {
    if (!p.k || p.k !== cfg.audiences.friend.token) return msgPage_('這個連結已經失效，請跟 ' + cfg.me.name + ' 要新的連結。');
    var bf = bookBoot_(cfg, 'friend', p.k); bf.here = selfUrl_(cfg, e);
    return page_('Book', bf, '跟 ' + cfg.me.name + ' 約時間', false);
  }
  if (p.v === 's') {
    if (!p.k || p.k !== cfg.shareToken) return msgPage_('這個連結不正確。');
    var bs = shareBoot_(cfg); bs.here = selfUrl_(cfg, e);
    return page_('Share', bs, cfg.me.name + ' 的行程', false);
  }
  var bp = bookBoot_(cfg, 'public', ''); bp.here = selfUrl_(cfg, e);
  return page_('Book', bp, '新朋友預約' + (cfg.me.brand ? '｜' + cfg.me.brand : ''), false);
}

/** 這一頁自己的完整網址（頁面在 Google 的 iframe 裡，抓不到真正的網址，所以由後端給） */
function selfUrl_(cfg, e) {
  var base = (cfg && cfg.publicUrl) || ScriptApp.getService().getUrl();
  var p = (e && e.parameter) || {}, qs = [];
  ['admin', 'v', 'k'].forEach(function (n) { if (p[n] !== undefined) qs.push(n + (p[n] === '' ? '' : '=' + encodeURIComponent(p[n]))); });
  return base + (qs.length ? '?' + qs.join('&') : '');
}

function page_(name, boot, title, isAdmin) {
  var t = HtmlService.createTemplateFromFile(name);
  t.boot = JSON.stringify(boot).replace(/</g, '\\u003c');
  var out = t.evaluate().setTitle(title)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    .addMetaTag('apple-mobile-web-app-capable', 'yes')
    .addMetaTag('mobile-web-app-capable', 'yes')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);   // 手機 App 內建瀏覽器才不會擋掉整頁
  return out;
}

function msgPage_(msg) {
  return HtmlService.createHtmlOutput('<div style="font-family:sans-serif;padding:40px 20px;text-align:center;color:#1d1c1a;min-height:100vh">' + esc_(msg) + '</div>')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1').setTitle('預約')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(name) { return HtmlService.createHtmlOutputFromFile(name).getContent(); }

/** 第一個月的資料直接放進頁面，打開就看得到，不用再等一次連線 */
function bookBoot_(cfg, aud, k) {
  var today = ymdOf_(new Date()), ym = today.slice(0, 7), first = null;
  try { var md = monthData_(cfg, aud, ym); first = { ym: ym, days: md.days, booked: md.booked }; } catch (e) { console.error(e); }
  return { aud: aud, k: k, today: today, rangeEnd: RANGE_END, first: first, notes: notesFor_(cfg, aud), me: cfg.me };
}
/** 每個月的話，挑出這個對象看得到的 */
function notesFor_(cfg, aud) {
  var out = {}, key = aud === 'friend' ? 'f' : 'p';
  Object.keys(cfg.notes || {}).forEach(function (ym) {
    var n = cfg.notes[ym];
    if (n && n[key]) out[ym] = n[key];
  });
  return out;
}
function shareBoot_(cfg) {
  var today = ymdOf_(new Date()), ym = today.slice(0, 7), first = null;
  try { first = { ym: ym, data: shareData_(cfg, ym) }; } catch (e) { console.error(e); }
  return { k: cfg.shareToken, today: today, rangeEnd: RANGE_END, first: first, me: cfg.me };
}

/* ======================= 權限 ======================= */

function isOwner_() {
  var a = Session.getActiveUser().getEmail();
  var me = Session.getEffectiveUser().getEmail();
  return !!a && !!me && a.toLowerCase() === me.toLowerCase();
}
function requireOwner_() { if (!isOwner_()) throw new Error('沒有權限'); }

function checkAud_(cfg, aud, k) {
  if (aud === 'friend') { if (!k || k !== cfg.audiences.friend.token) throw new Error('連結已失效'); }
  else if (aud !== 'public') throw new Error('連結錯誤');
}

/* ======================= 設定存取 ======================= */

function ss_() {
  var id = PropertiesService.getScriptProperties().getProperty('ssid');
  if (!id) throw new Error('尚未執行 setup');
  return SpreadsheetApp.openById(id);
}
function getConfig_() {
  var cache = CacheService.getScriptCache();
  var c = cache.get('cfg');
  if (c) return migrate_(JSON.parse(c));
  var v = ss_().getSheetByName(CFG_SHEET).getRange('A1').getValue();
  if (!v) throw new Error('尚未執行 setup');
  cache.put('cfg', v, 300);
  return migrate_(JSON.parse(v));
}
function migrate_(cfg) {
  if (!cfg.open) cfg.open = {};
  if (!cfg.people) cfg.people = [];
  if (!cfg.shareCals) cfg.shareCals = [];
  if (!cfg.weekHide) cfg.weekHide = [];
  if (!cfg.review) cfg.review = { friend: false, public: true };
  if (cfg.familyCal === undefined) { cfg.familyCal = ''; cfg.familyName = ''; }
  if (!cfg.me) cfg.me = { name: '我', brand: '', tagline: '' };
  if (!cfg.notes) cfg.notes = {};
  Object.keys(cfg.notes).forEach(function (ym) {          // 舊格式 { t, aud } 轉成朋友／新朋友分開
    var n = cfg.notes[ym];
    if (n && n.t !== undefined) cfg.notes[ym] = { f: (!n.aud || n.aud === 'all' || n.aud === 'friend') ? n.t : '', p: (!n.aud || n.aud === 'all' || n.aud === 'public') ? n.t : '' };
  });
  ['friend', 'public'].forEach(function (a) { delete cfg.audiences[a].durations; });
  if (!cfg.shareToken) { cfg.shareToken = newToken_(); saveConfig_(cfg); }
  return cfg;
}
function saveConfig_(cfg) {
  var s = JSON.stringify(cfg);
  ss_().getSheetByName(CFG_SHEET).getRange('A1').setValue(s);
  CacheService.getScriptCache().put('cfg', s, 300);
}

/* ======================= 日期工具 ======================= */

function pad_(n) { return (n < 10 ? '0' : '') + n; }
function ymdOf_(d) { return Utilities.formatDate(d, TZ, 'yyyy-MM-dd'); }
function dateOf_(ymd) { var a = ymd.split('-').map(Number); return new Date(a[0], a[1] - 1, a[2]); }
function addDaysYmd_(ymd, n) { var d = dateOf_(ymd); d.setDate(d.getDate() + n); return ymdOf_(d); }
function minYmd_(a, b) { return a < b ? a : b; }
function toMin_(hhmm) { var a = String(hhmm).split(':'); return (+a[0]) * 60 + (+a[1] || 0); }
function hhmm_(min) { return pad_(Math.floor(min / 60)) + ':' + pad_(min % 60); }
function isoLocal_(ymd, min) {
  if (min >= 1440) { ymd = addDaysYmd_(ymd, Math.floor(min / 1440)); min = min % 1440; }
  return ymd + 'T' + hhmm_(min) + ':00';
}
function whenText_(ymd, s, e) {
  var d = dateOf_(ymd);
  return (d.getMonth() + 1) + '/' + d.getDate() + '（' + WDN[d.getDay()] + '）' + hhmm_(s) + '–' + hhmm_(e);
}
function durText_(m) { return m % 60 === 0 ? (m / 60) + ' 小時' : (m < 60 ? m + ' 分' : Math.floor(m / 60) + ' 小時 ' + (m % 60) + ' 分'); }
function localParts_(dt) {  // Date → {ymd, min}
  return { ymd: ymdOf_(dt), min: (+Utilities.formatDate(dt, TZ, 'H')) * 60 + (+Utilities.formatDate(dt, TZ, 'm')) };
}

/* ======================= 空檔計算 ======================= */

function openRanges_(cfg, aud, ymd) {
  return (cfg.open[ymd] || []).filter(function (r) { return !r.aud || r.aud === 'all' || r.aud === aud; })
    .map(function (r) { return { s: toMin_(r.s), e: toMin_(r.e), loc: r.loc }; })
    .filter(function (r) { return r.e > r.s; });
}

function busy_(cfg, fromMs, toMs) {
  var ids = cfg.calendars.map(function (c) { return c.id; }).concat(cfg.busyAlso || []);
  var seen = {};
  var items = ids.filter(function (id) { if (seen[id]) return false; seen[id] = 1; return true; }).map(function (id) { return { id: id }; });
  var res = Calendar.Freebusy.query({
    timeMin: new Date(fromMs).toISOString(), timeMax: new Date(toMs).toISOString(), timeZone: TZ, items: items
  });
  var out = [];
  Object.keys(res.calendars || {}).forEach(function (id) {
    (res.calendars[id].busy || []).forEach(function (b) { out.push({ s: new Date(b.start).getTime(), e: new Date(b.end).getTime() }); });
  });
  return out;
}

/** 從區段扣掉忙碌時間。pieces/busy 皆為毫秒 {s,e} */
function subtract_(pieces, busy) {
  busy.forEach(function (b) {
    var next = [];
    pieces.forEach(function (p) {
      if (b.e <= p.s || b.s >= p.e) { next.push(p); return; }
      if (b.s > p.s) next.push({ s: p.s, e: b.s, loc: p.loc });
      if (b.e < p.e) next.push({ s: b.e, e: p.e, loc: p.loc });
    });
    pieces = next;
  });
  return pieces;
}

/** [s,e) 是否被連續的空檔蓋滿；回傳 'both' / 'online' / null */
function cover_(segs, s, e) {
  var t = s, loc = 'both';
  segs.slice().sort(function (a, b) { return a.s - b.s; }).forEach(function (p) {
    if (t < e && p.s <= t && p.e > t) { if (p.loc !== 'both') loc = 'online'; t = p.e; }
  });
  return t >= e ? loc : null;
}

/** 計算多天的狀態與空檔（純函式，方便測試） */
function computeDays_(cfg, aud, fromYmd, toYmd, busy, nowMs) {
  var a = cfg.audiences[aud];
  var today = ymdOf_(new Date(nowMs));
  var earliest = nowMs;                         // 不設提前限制，只擋掉已經過去的時間
  var minDur = STEP;
  var out = [];
  for (var ymd = fromYmd; ymd <= toYmd; ymd = addDaysYmd_(ymd, 1)) {
    var day = { date: ymd, state: '', segs: [] };
    if (ymd < today) day.state = 'past';
    else if (ymd > RANGE_END) day.state = 'none';
    else {
      var ranges = openRanges_(cfg, aud, ymd);
      if (!ranges.length) day.state = 'none';
      else {
        var base = dateOf_(ymd).getTime();
        var pieces = ranges.map(function (r) { return { s: Math.max(base + r.s * 60000, earliest), e: base + r.e * 60000, loc: r.loc }; })
          .filter(function (p) { return p.e > p.s; });
        pieces = subtract_(pieces, busy);
        day.segs = pieces.map(function (p) {
          var s = Math.round((p.s - base) / 60000), e = Math.round((p.e - base) / 60000);
          s = Math.ceil(s / STEP) * STEP;               // 對齊半點
          return { s: s, e: e, loc: p.loc };
        }).filter(function (p) { return p.e - p.s >= minDur; })
          .sort(function (x, y) { return x.s - y.s; });
        if (!day.segs.length) day.state = 'full';
        else day.state = day.segs.some(function (p) { return p.loc === 'both'; }) ? 'ok' : 'online';
      }
    }
    out.push(day);
  }
  return out;
}

/* ======================= 對外 API ======================= */

function apiMonth(aud, k, ym) {
  var cfg = getConfig_();
  checkAud_(cfg, aud, k);
  return monthData_(cfg, aud, ym);
}
function monthData_(cfg, aud, ym) {
  if (!/^\d{4}-\d{2}$/.test(ym)) throw new Error('月份錯誤');
  var from = ym + '-01';
  var d = dateOf_(from); d.setMonth(d.getMonth() + 1); d.setDate(0);
  var to = ymdOf_(d);
  var fromMs = dateOf_(from).getTime(), toMs = dateOf_(addDaysYmd_(to, 1)).getTime() + 3600000;
  var now = Date.now();
  var needBusy = Object.keys(cfg.open).some(function (d) { return d >= from && d <= to && openRanges_(cfg, aud, d).length; });
  var busy = needBusy ? busy_(cfg, fromMs, toMs) : [];
  var days = computeDays_(cfg, aud, from, to, busy, now);
  var bk = bookedIn_(cfg, from, to);
  days.forEach(function (d) { var t = bk.byDay[d.date]; if (t) { d.booked = t.length; d.taken = t; } });
  return { days: days, booked: bk.total };
}

/** 這段期間透過預約頁約的（含待審核），只回傳時間，不回傳是誰 */
function bookedIn_(cfg, from, to) {
  var ids = {}, byDay = {}, total = 0;
  [cfg.bookTo.friend, cfg.bookTo.public].forEach(function (id) { if (id) ids[id] = 1; });
  Object.keys(ids).forEach(function (cid) {
    try {
      var res = Calendar.Events.list(cid, { timeMin: dateOf_(from).toISOString(), timeMax: dateOf_(addDaysYmd_(to, 1)).toISOString(), singleEvents: true, maxResults: 500 });
      (res.items || []).forEach(function (ev) {
        if (ev.status === 'cancelled' || !ev.start.dateTime || !/^透過預約頁/.test(ev.description || '')) return;
        var a = localParts_(new Date(ev.start.dateTime)), b = localParts_(new Date(ev.end.dateTime));
        (byDay[a.ymd] = byDay[a.ymd] || []).push({ s: a.min, e: b.ymd === a.ymd ? b.min : 1440 });
        total++;
      });
    } catch (e) { console.error(e); }
  });
  Object.keys(byDay).forEach(function (k) { byDay[k].sort(function (x, y) { return x.s - y.s; }); });
  return { byDay: byDay, total: total };
}

function clean_(s, max) { return String(s == null ? '' : s).replace(/[\u0000-\u001f]/g, ' ').trim().slice(0, max); }
function mapUrl_(place) {
  if (!place) return '';
  if (/^https?:\/\/(www\.)?(google\.[a-z.]+\/maps|maps\.google\.|maps\.app\.goo\.gl|goo\.gl\/maps)/i.test(place)) return place;
  return 'https://www.google.com/maps/search/?api=1&query=' + encodeURIComponent(place);
}
function isEmail_(s) { return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(s); }

function apiBook(req) {
  var cfg = getConfig_();
  var aud = req.aud;
  checkAud_(cfg, aud, req.k);
  var a = cfg.audiences[aud];
  var name = clean_(req.name, 40), contact = clean_(req.contact, 100), purpose = clean_(req.purpose, 300), place = clean_(req.place, 500);
  var ymd = String(req.date), s = +req.start, e = +req.end, dur = e - s, locMode = req.locMode === 'inperson' ? 'inperson' : 'online';
  if (!name) throw new Error('請填你是誰');
  if (!purpose) throw new Error('請填要幹嘛');
  if (aud === 'public' && !isEmail_(contact)) throw new Error('請填正確的 Email');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(ymd) || !(s >= 0 && s < 1440) || !(e > s && e <= 1440) || s % STEP || e % STEP) throw new Error('時間錯誤');
  if (locMode === 'inperson' && !place) throw new Error('請填地點');
  if (/^https?:\/\//i.test(place) && !/^https?:\/\/(www\.)?(google\.[a-z.]+\/maps|maps\.google\.|maps\.app\.goo\.gl|goo\.gl\/maps)/i.test(place)) throw new Error('地點連結請貼 Google 地圖的連結');

  var lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) throw new Error('現在有點忙，請再按一次');
  try {
    var base = dateOf_(ymd).getTime();
    var busy = busy_(cfg, base - 86400000, base + 2 * 86400000);
    var day = computeDays_(cfg, aud, ymd, ymd, busy, Date.now())[0];
    var cov = cover_(day.segs || [], s, s + dur);
    if (!cov) throw new Error('這個時段剛被約走了，請換一個');
    if (locMode === 'inperson' && cov !== 'both') throw new Error('這個時段只開放線上');

    var who = AUD_NAME[aud];
    var review = !!(cfg.review && cfg.review[aud]);          // 要先經過你同意
    var calId = cfg.bookTo[aud];
    var locText = locMode === 'online' ? '線上（Google Meet）' : place;
    var map = locMode === 'online' ? '' : mapUrl_(place);
    var desc = bookDesc_(who, name, contact, purpose, durText_(dur), locText, map, review);
    var ev = insertEvent_(calId, {
      title: (review ? '【待審核】' : '【' + who + '】') + name + '｜' + purpose.slice(0, 30),
      ymd: ymd, s: s, e: s + dur, loc: locText, desc: desc,
      attendees: review ? [] : (isEmail_(contact) ? [contact] : []), meet: locMode === 'online' && !review, who: name
    });
    var meet = meetLink_(ev);
    var when = whenText_(ymd, s, s + dur);
    appendInbox_({ status: review ? '待審核' : '已排', type: '指定時段', aud: who, name: name, contact: contact, when: when, dur: durText_(dur),
      purpose: purpose, locMode: locMode === 'online' ? '線上' : '實體', place: locMode === 'online' ? '' : place, meet: meet, calId: calId, eventId: ev.id });
    notify_(cfg, (review ? '待你確認' : '新預約') + '｜' + who + '｜' + name + '｜' + when, [
      ['誰', name + (contact ? '（' + contact + '）' : '')], ['要幹嘛', purpose], ['多久', durText_(dur)], ['地點', /^https?:/i.test(locText) ? 'Google 地圖連結（點下方按鈕）' : locText],
      ['行事曆', calName_(cfg, calId) + (review ? '（已先卡住這個時段）' : '（要取消直接刪行程）')]
    ], { tag: review ? '待你確認' : who + '預約', heading: name, when: when,
      buttons: [review && adminLink_(cfg) && ['去後台通過或婉拒', adminLink_(cfg)], meet && ['加入 Google Meet', meet], map && ['在 Google 地圖打開', map], ev.htmlLink && ['在 Google 日曆打開', ev.htmlLink]] });
    return { ok: true, pending: review, when: when, loc: locText, meet: meet, purpose: purpose, map: map };
  } finally {
    lock.releaseLock();
  }
}

function bookDesc_(who, name, contact, purpose, dur, locText, map, review) {
  return ['透過預約頁（' + who + '）', '誰：' + name, '聯絡：' + (contact || '—'), '要幹嘛：' + purpose, '多久：' + dur, '約哪：' + locText]
    .concat(map ? ['地圖：' + map] : []).concat(review ? ['狀態：等本人確認'] : []).join('\n');
}
function adminLink_(cfg) { var b = cfg.publicUrl || ScriptApp.getService().getUrl(); return b ? b + '?admin' : ''; }
function bookLink_(cfg) { return cfg.shortPublic || (cfg.publicUrl || ScriptApp.getService().getUrl()); }

function apiRequestMonth(req) {
  var cfg = getConfig_();
  var aud = req.aud;
  checkAud_(cfg, aud, req.k);
  var name = clean_(req.name, 40), contact = clean_(req.contact, 100), purpose = clean_(req.purpose, 300);
  var month = String(req.month), dur = clean_(req.dur, 20), locPref = clean_(req.locPref, 20);
  if (!name) throw new Error('請填你是誰');
  if (!purpose) throw new Error('請填要幹嘛');
  if (aud === 'public' && !isEmail_(contact)) throw new Error('請填正確的 Email');
  if (!/^\d{4}-\d{2}$/.test(month) || month > RANGE_END.slice(0, 7)) throw new Error('月份錯誤');
  var who = AUD_NAME[aud];
  var mt = month.slice(0, 4) + ' 年 ' + (+month.slice(5)) + ' 月';
  appendInbox_({ status: '待處理', type: '只約月份', aud: who, name: name, contact: contact, when: mt, dur: dur,
    purpose: purpose, locMode: locPref, place: '', meet: '', calId: '', eventId: '' });
  notify_(cfg, '新留言｜只約月份｜' + who + '｜' + name + '｜' + mt, [
    ['誰', name + (contact ? '（' + contact + '）' : '')], ['要幹嘛', purpose], ['多久', dur || '—'], ['約哪', locPref || '—'],
    ['下一步', '到後台收件匣按「排時間」']
  ], { tag: who + '・只約月份', heading: name, when: mt });
  return { ok: true, month: mt };
}

/** 分享連結：只回傳勾選的行事曆，只有標題和時間；預約進來的行程不顯示對方資料 */
function apiShare(k, ym) {
  var cfg = getConfig_();
  if (!k || k !== cfg.shareToken) throw new Error('連結不正確');
  return shareData_(cfg, ym);
}
/** 分享頁一個月的行程；快取 1 分鐘，讓分享頁開得快 */
function shareData_(cfg, ym) {
  if (!/^\d{4}-\d{2}$/.test(ym)) throw new Error('月份錯誤');
  var ck = 'sh_' + ym + '_' + (cfg.shareVer || 0), cache = CacheService.getScriptCache();
  try { var hit = cache.get(ck); if (hit) return JSON.parse(hit); } catch (e) {}
  var r = shareFetch_(cfg, ym);
  try { var js = JSON.stringify(r); if (js.length < 90000) cache.put(ck, js, 60); } catch (e) {}
  return r;
}
function shareFetch_(cfg, ym) {
  var from = ym + '-01', d = dateOf_(from); d.setMonth(d.getMonth() + 1);
  var timeMin = dateOf_(from).toISOString(), timeMax = d.toISOString();
  var shared = {}, colorOf = {}, out = [], cals = [];
  cfg.calendars.forEach(function (c) { colorOf[c.id] = c.color; if (cfg.shareCals.indexOf(c.id) >= 0) { shared[c.id] = 1; cals.push({ name: c.name, color: c.color }); } });
  if (!cals.length) return { events: [], calendars: [] };
  cfg.calendars.forEach(function (c) {
    var token = null;
    do {
      var opt = { timeMin: timeMin, timeMax: timeMax, singleEvents: true, orderBy: 'startTime', maxResults: 2500 };
      if (token) opt.pageToken = token;
      var res = Calendar.Events.list(c.id, opt);
      (res.items || []).forEach(function (ev) {
        if (ev.status === 'cancelled') return;
        var pv = (ev.extendedProperties && ev.extendedProperties.private) || {};
        var mine = [c.id].concat(tagsOf_(cfg, pv.tags, c.id)).filter(function (id) { return shared[id]; });
        if (!mine.length) return;                                   // 沒有任何一本有分享
        var booked = /^透過預約頁/.test(ev.description || '');
        var o = { title: booked ? '已有預約' : (ev.summary || '（無標題）'), color: colorOf[mine[0]] };
        if (ev.start.date) { o.allDay = true; o.sd = ev.start.date; o.ed = addDaysYmd_(ev.end.date, -1); o.sm = 0; o.em = 1440; }
        else {
          var a = localParts_(new Date(ev.start.dateTime)), b = localParts_(new Date(ev.end.dateTime));
          o.allDay = false; o.sd = a.ymd; o.sm = a.min; o.ed = b.ymd; o.em = b.min;
        }
        out.push(o);
      });
      token = res.nextPageToken;
    } while (token);
  });
  out.sort(function (x, y) { return (x.sd + (x.allDay ? '0' : '1') + ('000' + x.sm).slice(-4)) < (y.sd + (y.allDay ? '0' : '1') + ('000' + y.sm).slice(-4)) ? -1 : 1; });
  return { events: out, calendars: cals };
}

/* ======================= 行事曆事件 ======================= */

function insertEvent_(calId, o) {
  var ev = { summary: o.title, location: o.loc || '', description: o.desc || o.note || '' };
  var tags = (o.tags || []).join(',');
  if (o.who || o.note || tags || o.famId) ev.extendedProperties = { private: { who: o.who || '', note: o.note || '', tags: tags, famId: o.famId || '' } };
  if (o.allDay) {
    ev.start = { date: o.ymd };
    ev.end = { date: addDaysYmd_(o.endYmd || o.ymd, 1) };
    ev.transparency = 'opaque';
  } else {
    ev.start = { dateTime: isoLocal_(o.ymd, o.s), timeZone: TZ };
    ev.end = { dateTime: isoLocal_(o.endYmd || o.ymd, o.e), timeZone: TZ };
  }
  if (o.repeat && o.repeat.length) {
    var until = (o.until && o.until <= RANGE_END ? o.until : RANGE_END).replace(/-/g, '');
    ev.recurrence = ['RRULE:FREQ=WEEKLY;BYDAY=' + o.repeat.map(function (d) { return BYDAY[d]; }).join(',') + ';UNTIL=' + until + 'T155959Z'];
    var ex = (o.exdates || []).filter(function (d) { return /^\d{4}-\d{2}-\d{2}$/.test(d); });
    if (ex.length) ev.recurrence.push(o.allDay ? 'EXDATE;VALUE=DATE:' + ex.map(function (d) { return d.replace(/-/g, ''); }).join(',')
      : 'EXDATE;TZID=' + TZ + ':' + ex.map(function (d) { return d.replace(/-/g, '') + 'T' + hhmm_(o.s).replace(':', '') + '00'; }).join(','));
  }
  var opt = {};
  if (o.attendees && o.attendees.length) { ev.attendees = o.attendees.map(function (e) { return { email: e }; }); opt.sendUpdates = 'all'; }
  if (o.meet) {
    ev.conferenceData = { createRequest: { requestId: Utilities.getUuid(), conferenceSolutionKey: { type: 'hangoutsMeet' } } };
    opt.conferenceDataVersion = 1;
  }
  return Calendar.Events.insert(ev, calId, opt);
}

function meetLink_(ev) {
  if (ev.hangoutLink) return ev.hangoutLink;
  var ep = ev.conferenceData && ev.conferenceData.entryPoints;
  if (ep) for (var i = 0; i < ep.length; i++) if (ep[i].entryPointType === 'video') return ep[i].uri;
  return '';
}

function calName_(cfg, id) {
  var c = cfg.calendars.filter(function (x) { return x.id === id; })[0];
  return c ? c.name : '';
}

/* ======================= 收件匣與通知 ======================= */

function appendInbox_(r) {
  var sh = ss_().getSheetByName(INBOX);
  sh.appendRow([Utilities.getUuid().slice(0, 8), Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm'), r.status, r.type, r.aud,
    r.name, r.contact, r.when, r.dur, r.purpose, r.locMode, r.place, r.meet, r.calId, r.eventId].map(function (v) {
      v = v == null ? '' : String(v);
      return /^[=+\-@]/.test(v) ? "'" + v : v;     // 防止被當成公式
    }));
}

/** 寄給自己的通知信：上方是誰和時間，中間是資料，下方是按鈕 */
function notify_(cfg, subject, rows, o) {
  sendCard_(cfg.notifyEmail || Session.getEffectiveUser().getEmail(), subject, rows, o);
}

/** 寄一封卡片式的信 */
function sendCard_(to, subject, rows, o) {
  o = o || {};
  var btns = (o.buttons || []).filter(Boolean);
  var html = '<div style="background:#f6f5f2;padding:24px 12px;font-family:-apple-system,BlinkMacSystemFont,\'PingFang TC\',\'Noto Sans TC\',sans-serif;color:#1d1c1a">' +
    '<div style="max-width:480px;margin:0 auto;background:#fff;border-radius:16px;overflow:hidden">' +
    '<div style="padding:20px 22px 16px;border-bottom:1px solid #efece6">' +
    (o.tag ? '<div style="display:inline-block;font-size:12px;font-weight:700;color:#a95a20;background:#fbeee4;border-radius:6px;padding:3px 8px">' + esc_(o.tag) + '</div>' : '') +
    '<div style="font-size:22px;font-weight:700;margin-top:10px">' + esc_(o.heading || subject) + '</div>' +
    (o.when ? '<div style="font-size:16px;margin-top:4px;color:#1d1c1a">' + esc_(o.when) + '</div>' : '') + '</div>' +
    '<div style="padding:10px 22px 6px">' + rows.map(function (r) {
      return '<div style="display:flex;padding:9px 0;border-bottom:1px solid #f3f1ec;font-size:14px;line-height:1.5"><div style="width:64px;flex-shrink:0;color:#77736c">' + esc_(r[0]) + '</div><div style="flex:1">' + esc_(r[1]) + '</div></div>';
    }).join('') + '</div>' +
    (btns.length ? '<div style="padding:10px 22px 22px">' + btns.map(function (b, i) {
      return '<a href="' + esc_(b[1]) + '" style="display:block;text-align:center;text-decoration:none;font-weight:700;font-size:14px;border-radius:10px;padding:12px;margin-top:8px;' +
        (i === 0 ? 'background:#1d1c1a;color:#fff' : 'background:#f3f1ec;color:#1d1c1a') + '">' + esc_(b[0]) + '</a>';
    }).join('') + '</div>' : '') +
    (o.foot ? '<div style="padding:0 22px 20px;font-size:13px;color:#77736c;line-height:1.6">' + esc_(o.foot) + '</div>' : '') +
    '</div><div style="max-width:480px;margin:10px auto 0;font-size:11px;color:#9a968e;text-align:center">預約系統自動寄出</div></div>';
  var text = [o.heading, o.when].filter(Boolean).join('｜') + '\n' + rows.map(function (r) { return r[0] + '：' + r[1]; }).join('\n') +
    btns.map(function (b) { return '\n' + b[0] + '：' + b[1]; }).join('') + (o.foot ? '\n\n' + o.foot : '');
  try { MailApp.sendEmail({ to: to, subject: subject, body: text, htmlBody: html, name: o.from || '預約系統' }); } catch (e) { console.error(e); }
}

function esc_(s) { return String(s).replace(/[&<>"']/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]; }); }

/* ======================= 後台 API（只有本人） ======================= */

function adminBoot() {
  requireOwner_();
  var cfg = getConfig_();
  var base = cfg.publicUrl || ScriptApp.getService().getUrl();
  return {
    cfg: cfg,
    friendUrl: base + '?v=f&k=' + cfg.audiences.friend.token + '&openExternalBrowser=1',   // LINE 會改用手機瀏覽器開
    shareUrl: base + '?v=s&k=' + cfg.shareToken + '&openExternalBrowser=1',
    publicUrl: base + '?openExternalBrowser=1',
    sheetUrl: ss_().getUrl(),
    today: ymdOf_(new Date()),
    rangeEnd: RANGE_END,
    board: getBoard_(),
    ai: aiStatus_()
  };
}

function adminEvents(fromYmd, toYmd) {
  requireOwner_();
  var cfg = getConfig_();
  var timeMin = dateOf_(fromYmd).toISOString(), timeMax = dateOf_(addDaysYmd_(toYmd, 1)).toISOString();
  var out = [];
  cfg.calendars.forEach(function (c) {
    var token = null;
    do {
      var opt = { timeMin: timeMin, timeMax: timeMax, singleEvents: true, orderBy: 'startTime', maxResults: 2500 };
      if (token) opt.pageToken = token;
      var res = Calendar.Events.list(c.id, opt);
      (res.items || []).forEach(function (ev) {
        if (ev.status === 'cancelled') return;
        var o = { cal: c.id, id: ev.id, rid: ev.recurringEventId || '', title: ev.summary || '（無標題）', loc: ev.location || '',
          desc: (ev.description || '').slice(0, 500), meet: meetLink_(ev), booked: /^透過預約頁/.test(ev.description || ''),
          who: (ev.extendedProperties && ev.extendedProperties.private && ev.extendedProperties.private.who) || '' };
        var pv = (ev.extendedProperties && ev.extendedProperties.private) || {};
        o.tags = tagsOf_(cfg, pv.tags, c.id);
        o.famId = pv.famId || '';
        o.note = pv.note != null ? String(pv.note) : (o.booked ? '' : (ev.description || '').slice(0, 2000));   // 自己的行程：備註就是說明欄
        if (ev.start.date) {
          o.allDay = true; o.sd = ev.start.date; o.ed = addDaysYmd_(ev.end.date, -1); o.sm = 0; o.em = 1440;
        } else {
          var a = localParts_(new Date(ev.start.dateTime)), b = localParts_(new Date(ev.end.dateTime));
          o.allDay = false; o.sd = a.ymd; o.sm = a.min; o.ed = b.ymd; o.em = b.min;
        }
        out.push(o);
      });
      token = res.nextPageToken;
    } while (token);
  });
  out.sort(function (x, y) { return (x.sd + pad_(x.allDay ? 0 : 1) + ('000' + x.sm).slice(-4)) < (y.sd + pad_(y.allDay ? 0 : 1) + ('000' + y.sm).slice(-4)) ? -1 : 1; });
  return out;
}

/** 同一筆行程另外歸到哪些行事曆（去掉已刪的、跟主要重複的） */
function tagsOf_(cfg, raw, main) {
  var ok = {};
  cfg.calendars.forEach(function (c) { ok[c.id] = 1; });
  var seen = {}, list = Array.isArray(raw) ? raw : String(raw || '').split(',');
  return list.map(function (x) { return String(x).trim(); }).filter(function (x) {
    if (!x || !ok[x] || x === main || seen[x]) return false; seen[x] = 1; return true;
  });
}
function validCal_(cfg, id) { if (!cfg.calendars.some(function (c) { return c.id === id; })) throw new Error('找不到這個行事曆'); }

function adminCreate(o) {
  requireOwner_();
  var cfg = getConfig_();
  validCal_(cfg, o.cal);
  if (!o.title) throw new Error('請填內容');
  var who = clean_(o.who, 40);
  var famId = o.family ? famWrite_(cfg, o, '') : '';
  var ev = insertEvent_(o.cal, { title: clean_(o.title, 200), ymd: o.date, endYmd: o.endDate, s: +o.s, e: +o.e, allDay: !!o.allDay,
    loc: clean_(o.loc, 500), repeat: o.repeat, until: o.until, exdates: o.exdates, meet: !!o.meet && !o.allDay, who: who, note: note_(o.note), tags: tagsOf_(cfg, o.tags, o.cal), famId: famId });
  if (rememberWho_(cfg, who)) saveConfig_(cfg);
  return { id: ev.id, people: cfg.people, meet: meetLink_(ev) };
}

/** 整批匯入：跳過同一天、同時間、同標題已經存在的行程 */
function adminImport(items) {
  requireOwner_();
  var cfg = getConfig_();
  items = (items || []).filter(function (o) { return o && o.title && /^\d{4}-\d{2}-\d{2}$/.test(o.date); });
  if (!items.length) throw new Error('沒有可以匯入的行程');
  if (items.length > 300) throw new Error('一次最多 300 筆');
  var from = items.reduce(function (m, o) { return o.date < m ? o.date : m; }, items[0].date);
  var to = items.reduce(function (m, o) { var e = o.until || o.endDate || o.date; return e > m ? e : m; }, items[0].date);
  var seen = {};
  cfg.calendars.forEach(function (c) {
    var token = null;
    do {
      var opt = { timeMin: dateOf_(from).toISOString(), timeMax: dateOf_(addDaysYmd_(to, 1)).toISOString(), singleEvents: true, maxResults: 2500 };
      if (token) opt.pageToken = token;
      var res = Calendar.Events.list(c.id, opt);
      (res.items || []).forEach(function (ev) {
        if (ev.status === 'cancelled') return;
        var st = ev.start.date ? ev.start.date + '|all' : (function (p) { return p.ymd + '|' + p.min; })(localParts_(new Date(ev.start.dateTime)));
        seen[st + '|' + (ev.summary || '')] = 1;
      });
      token = res.nextPageToken;
    } while (token);
  });
  var added = 0, skipped = 0;
  items.forEach(function (o) {
    validCal_(cfg, o.cal);
    var key = o.date + '|' + (o.allDay ? 'all' : +o.s) + '|' + clean_(o.title, 200);
    if (seen[key]) { skipped++; return; }
    insertEvent_(o.cal, { title: clean_(o.title, 200), ymd: o.date, endYmd: o.endDate, s: +o.s, e: +o.e, allDay: !!o.allDay,
      loc: clean_(o.loc, 500), repeat: o.repeat, until: o.until, exdates: o.exdates, who: clean_(o.who, 40), note: note_(o.note), tags: tagsOf_(cfg, o.tags, o.cal) });
    rememberWho_(cfg, clean_(o.who, 40));
    seen[key] = 1; added++;
  });
  saveConfig_(cfg);
  return { added: added, skipped: skipped };
}

function adminUpdate(o) {
  requireOwner_();
  var cfg = getConfig_();
  validCal_(cfg, o.cal);
  var who = clean_(o.who, 40);
  var note = note_(o.note);
  var famId = o.famId || '';
  if (!o.rid) {                                   // 重複行程的單一次不同步家庭日曆（整串在新增時就決定）
    if (o.family) famId = famWrite_(cfg, o, famId);
    else if (famId) { famRemove_(cfg, famId); famId = ''; }
  }
  var patch = { summary: clean_(o.title, 200), location: clean_(o.loc, 500),
    extendedProperties: { private: { who: who, note: note, tags: tagsOf_(cfg, o.tags, o.newCal || o.cal).join(','), famId: famId } } };
  if (!o.booked) patch.description = note;       // 預約來的行程保留原本的預約資料
  if (rememberWho_(cfg, who)) saveConfig_(cfg);
  if (o.allDay) { patch.start = { date: o.date, dateTime: null }; patch.end = { date: addDaysYmd_(o.endDate || o.date, 1), dateTime: null }; }
  else { patch.start = { dateTime: isoLocal_(o.date, +o.s), timeZone: TZ, date: null }; patch.end = { dateTime: isoLocal_(o.endDate || o.date, +o.e), timeZone: TZ, date: null }; }
  var popt = {};
  if (o.addMeet && !o.allDay) { patch.conferenceData = { createRequest: { requestId: Utilities.getUuid(), conferenceSolutionKey: { type: 'hangoutsMeet' } } }; popt.conferenceDataVersion = 1; }
  Calendar.Events.patch(patch, o.cal, o.id, popt);
  if (o.newCal && o.newCal !== o.cal) {
    validCal_(cfg, o.newCal);
    Calendar.Events.move(o.cal, o.rid || o.id, o.newCal);   // 重複行程會整串一起移
  }
  return { ok: true, people: cfg.people };
}

function adminDelete(cal, id, rid, all) {
  requireOwner_();
  var cfg = getConfig_();
  validCal_(cfg, cal);
  var target = all && rid ? rid : id;
  if (!rid || all) {                              // 家庭日曆那份一起刪
    try { var ev = Calendar.Events.get(cal, target); var pv = (ev.extendedProperties && ev.extendedProperties.private) || {}; if (pv.famId) famRemove_(cfg, pv.famId); } catch (e) { console.error(e); }
  }
  Calendar.Events.remove(cal, target);
  return { ok: true };
}

/* ======================= 截圖判讀：Claude（付費、快）或 Gemini（免費、較慢） ======================= */
/* 金鑰存在「指令碼屬性」，不在試算表、也不在程式碼裡。貼哪一家的金鑰，就用哪一家 */

var CLAUDE_MODELS = ['claude-haiku-4-5-20251001', 'claude-haiku-4-5'];   // 最便宜的 Claude
var AI_MODELS = ['gemini-flash-latest', 'gemini-3.8-flash', 'gemini-3.5-flash', 'gemini-flash-lite-latest', 'gemini-3.5-flash-lite'];
function aiKey_() {
  var p = PropertiesService.getScriptProperties();
  var k = p.getProperty('AI_KEY') || p.getProperty('GEMINI_KEY') || '';
  return { key: k, provider: /^sk-ant-/.test(k) ? 'claude' : 'gemini' };
}
function aiStatus_() {
  var a = aiKey_();
  return { set: !!a.key, tail: a.key ? a.key.slice(-4) : '', provider: a.key ? a.provider : '' };
}
function adminSaveAiKey(key) {
  requireOwner_();
  var props = PropertiesService.getScriptProperties();
  key = String(key || '').replace(/\s+/g, '');                     // 複製時多帶的空白、換行去掉
  if (!key) { props.deleteProperty('AI_KEY'); props.deleteProperty('GEMINI_KEY'); return aiStatus_(); }
  if (!/^[A-Za-z0-9_\-.]{20,300}$/.test(key)) throw new Error('金鑰格式不對，請整串複製貼上（Claude 是 sk-ant- 開頭，Gemini 是 AQ. 或 AIza 開頭）');
  props.setProperty('AI_KEY', key); props.deleteProperty('GEMINI_KEY');
  return aiStatus_();
}

function aiPrompt_() {
  var now = new Date(), today = ymdOf_(now);
  return [
    '你是行程助理。圖片是一段聊天截圖（LINE、IG、Messenger 之類）。',
    '找出裡面已經約定好、或正在提議的見面與活動，只輸出一個 JSON 陣列，不要輸出任何其他文字或說明。',
    '今天是 ' + today + '（星期' + WDN[now.getDay()] + '），時區台北。',
    '「明天」「下週四」「這週六」這類相對日期要換算成實際日期；沒寫年份就用今天之後最近的那個日期。',
    '每一筆格式：{"date":"YYYY-MM-DD","start":"HH:MM 或 null","end":"HH:MM 或 null","allDay":true 或 false,',
    '"title":"簡短標題，例如「跟阿明吃飯」","who":"對方的名字或暱稱，看不出來就空字串","place":"地點，沒有就空字串",',
    '"note":"其他重要細節，例如要帶的東西，沒有就空字串","sure":true 或 false（日期時間是否已經講定）}',
    '只有時間點沒有結束時間時 end 給 null。完全沒有活動就輸出 []。'
  ].join('\n');
}

/** 讀一張截圖，回傳裡面約好的行程（還沒寫進日曆，後台確認後才匯入） */
function adminParseImage(b64, mime) {
  requireOwner_();
  var a = aiKey_();
  if (!a.key) throw new Error('還沒設定截圖判讀的金鑰（設定 → 截圖判讀）');
  if (!/^image\/(png|jpeg|webp)$/.test(mime)) throw new Error('只支援 PNG、JPG 截圖');
  if (!b64 || b64.length > 8000000) throw new Error('圖片太大，請截小一點');
  return cleanAiEvents_(a.provider === 'claude' ? claudeRead_(a.key, b64, mime) : geminiRead_(a.key, b64, mime));
}

/** Claude（Messages API） */
function claudeRead_(key, b64, mime) {
  var body = { max_tokens: 1500, temperature: 0,
    messages: [{ role: 'user', content: [{ type: 'image', source: { type: 'base64', media_type: mime, data: b64 } }, { type: 'text', text: aiPrompt_() }] }] };
  var last = 0, started = Date.now();
  for (var i = 0; i < CLAUDE_MODELS.length; i++) {
    body.model = CLAUDE_MODELS[i];
    for (var attempt = 0; attempt < 3; attempt++) {
      if (Date.now() - started > 45000) break;
      var res = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
        method: 'post', contentType: 'application/json', muteHttpExceptions: true, payload: JSON.stringify(body),
        headers: { 'x-api-key': key, 'anthropic-version': '2023-06-01' } });
      var code = res.getResponseCode(), txt = res.getContentText();
      last = code;
      if (code === 200) {
        try { return JSON.parse(txt).content.map(function (c) { return c.text || ''; }).join(''); } catch (e) { return ''; }
      }
      if (code === 401) throw new Error('Claude 金鑰無效（401），請到設定重新貼一次');
      if (code === 403) throw new Error('這組 Claude 金鑰沒有權限（403），請到 console.anthropic.com 確認');
      if (code === 400 && /credit balance/i.test(txt)) throw new Error('Claude 帳戶餘額不足，請到 console.anthropic.com → Billing 儲值');
      if (code === 429 || code === 500 || code === 529 || code === 503) { Utilities.sleep(1500 * (attempt + 1)); continue; }   // 太忙或太頻繁：等一下再試
      break;                                                       // 404 型號名稱不對 → 換下一個
    }
  }
  if (last === 429) throw new Error('Claude 暫時請求太多（429），過一分鐘再試');
  if (last >= 500) throw new Error('Claude 現在太忙（' + last + '），過幾分鐘再按一次「判讀」');
  throw new Error('判讀失敗（' + last + '），稍後再試');
}

/** Gemini（免費額度） */
function geminiRead_(key, b64, mime) {
  var body = { contents: [{ parts: [{ text: aiPrompt_() }, { inline_data: { mime_type: mime, data: b64 } }] }],
    generationConfig: { responseMimeType: 'application/json', temperature: 0.1 } };
  var cfg = getConfig_(), tried = [cfg.aiModel].concat(AI_MODELS).filter(function (m, i, a) { return m && /^gemini/.test(m) && a.indexOf(m) === i; });
  var last = 0, started = Date.now();
  for (var i = 0; i < tried.length; i++) {
    for (var attempt = 0; attempt < 3; attempt++) {
      if (Date.now() - started > 45000) break;                   // 最多等 45 秒，不要讓畫面卡太久
      var res = UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models/' + tried[i] + ':generateContent', {
        method: 'post', contentType: 'application/json', headers: { 'x-goog-api-key': key }, payload: JSON.stringify(body), muteHttpExceptions: true });
      var code = res.getResponseCode(), txt = res.getContentText();
      last = code;
      if (code === 200) {
        if (cfg.aiModel !== tried[i]) { cfg.aiModel = tried[i]; saveConfig_(cfg); }
        try { return JSON.parse(txt).candidates[0].content.parts.map(function (p) { return p.text || ''; }).join(''); } catch (e) { return ''; }
      }
      if (code === 400 && /API key/i.test(txt)) throw new Error('Gemini 金鑰無效，請到設定重新貼一次');
      if (code === 401) throw new Error('Gemini 不接受這組金鑰（401）。新建立的 AQ. 金鑰有時要等幾分鐘才生效；還是不行就到 AI Studio 重新建立一組');
      if (code === 403) throw new Error('這組金鑰沒有權限，請確認是在 Google AI Studio 建立的');
      if (code === 500 || code === 502 || code === 503 || code === 504) { Utilities.sleep(1500 * (attempt + 1)); continue; }   // Gemini 太忙：等一下再試
      break;                                                     // 404 型號不存在、429 這個型號額度用完 → 換下一個型號
    }
  }
  if (last === 429) throw new Error('今天的免費額度用完了，明天再試或改用手動輸入');
  if (last >= 500) throw new Error('Gemini 現在太忙（' + last + '），過幾分鐘再按一次「判讀」');
  throw new Error('判讀失敗（' + last + '），稍後再試');
}
function cleanAiEvents_(text) {
  var list, raw = String(text || '').replace(/```(json)?/g, '').trim();
  var i = raw.indexOf('['), j = raw.lastIndexOf(']');              // 模型偶爾會在 JSON 前後多講一句話
  if (i >= 0 && j > i) raw = raw.slice(i, j + 1);
  try { list = JSON.parse(raw); } catch (e) { throw new Error('看不懂這張截圖，換一張或改用手動輸入'); }
  if (!Array.isArray(list)) list = list && Array.isArray(list.events) ? list.events : [];
  var hm = function (x) { var m = /^(\d{1,2}):(\d{2})$/.exec(String(x || '')); return m && +m[1] < 24 && +m[2] < 60 ? (+m[1]) * 60 + (+m[2]) : null; };
  return list.slice(0, 20).map(function (x) {
    if (!x || !/^\d{4}-\d{2}-\d{2}$/.test(x.date)) return null;
    var s = hm(x.start), e = hm(x.end), allDay = !!x.allDay || s == null;
    if (!allDay && (e == null || e <= s)) e = Math.min(s + 60, 1440);
    return { date: x.date, s: allDay ? 0 : s, e: allDay ? 1440 : e, allDay: allDay, title: clean_(x.title, 100) || '（未命名）',
      who: clean_(x.who, 40), loc: clean_(x.place, 200), note: note_(x.note).slice(0, 300), sure: x.sure !== false };
  }).filter(Boolean);
}

/* ======================= 家庭日曆：只寫入你選的行程，不讀家人的 ======================= */

/** 建立或更新家庭日曆裡的那一份（只放標題、時間、地點），回傳它的 id */
function famWrite_(cfg, o, famId) {
  if (!cfg.familyCal) throw new Error('還沒設定家庭日曆（設定 → 家庭日曆）');
  var base = { title: clean_(o.title, 200), ymd: o.date, endYmd: o.endDate, s: +o.s, e: +o.e, allDay: !!o.allDay, loc: clean_(o.loc, 500) };
  if (famId) {
    try {
      var p = { summary: base.title, location: base.loc };
      if (o.allDay) { p.start = { date: o.date, dateTime: null }; p.end = { date: addDaysYmd_(o.endDate || o.date, 1), dateTime: null }; }
      else { p.start = { dateTime: isoLocal_(o.date, +o.s), timeZone: TZ, date: null }; p.end = { dateTime: isoLocal_(o.endDate || o.date, +o.e), timeZone: TZ, date: null }; }
      Calendar.Events.patch(p, cfg.familyCal, famId);
      return famId;
    } catch (e) { console.error(e); }             // 被家人刪掉了 → 重新建一份
  }
  base.repeat = o.repeat; base.until = o.until; base.exdates = o.exdates;
  return insertEvent_(cfg.familyCal, base).id;
}
function famRemove_(cfg, famId) {
  if (!cfg.familyCal || !famId) return;
  try { Calendar.Events.remove(cfg.familyCal, famId); } catch (e) { console.error(e); }
}
/** 列出可以選的日曆（你能編輯、又不是那幾本分類的），順便猜哪本是家庭 */
function adminFamilyOptions() {
  requireOwner_();
  var cfg = getConfig_(), mine = {};
  cfg.calendars.forEach(function (c) { mine[c.id] = 1; });
  var items = (Calendar.CalendarList.list({ minAccessRole: 'writer', maxResults: 250 }).items || [])
    .filter(function (c) { return !mine[c.id] && !c.primary; })
    .map(function (c) { return { id: c.id, name: c.summaryOverride || c.summary || c.id }; });
  var guess = items.filter(function (c) { return /家庭|家人|family/i.test(c.name); })[0];
  return { items: items, current: cfg.familyCal, guess: guess ? guess.id : '' };
}
function adminSaveFamilyCal(id) {
  requireOwner_();
  var cfg = getConfig_();
  if (id) {
    var hit = adminFamilyOptions().items.filter(function (c) { return c.id === id; })[0];
    if (!hit) throw new Error('找不到這本日曆，或你沒有編輯權限');
    cfg.familyCal = id; cfg.familyName = hit.name;
  } else { cfg.familyCal = ''; cfg.familyName = ''; }
  saveConfig_(cfg);
  return cfg;
}

var REL_WORDS = ['朋友', '家人', '同事', '同學', '學長', '學姐', '學弟', '學妹', '主管', '老闆', '客戶', '爸媽', '爸爸', '媽媽', '男友', '女友', '老婆', '老公'];
/** 備註可以換行，其他控制字元去掉 */
function note_(s) { return String(s == null ? '' : s).replace(/\r\n?/g, '\n').replace(/[\u0000-\u0009\u000b-\u001f]/g, ' ').trim().slice(0, 2000); }

/* ======================= 白板（不限時的記事） ======================= */

var BOARD_SHEET = '_白板';
function boardSheet_() {
  var ss = ss_(), sh = ss.getSheetByName(BOARD_SHEET);
  if (!sh) { sh = ss.insertSheet(BOARD_SHEET); sh.hideSheet(); }
  return sh;
}
/** 白板：固定 8 格，每格 { t: 文字, d: 做完了沒 } */
var BOARD_N = 8;
function blankBoard_() { var a = []; for (var i = 0; i < BOARD_N; i++) a.push({ t: '', d: false }); return a; }
function getBoard_() {
  var v = boardSheet_().getRange('A1').getValue(), rows = [];
  if (v) {
    try {
      var o = JSON.parse(v);
      if (Array.isArray(o)) rows = o.map(function (n) { return { t: String(n.t || ''), d: !!(n.d || n.done) }; });   // 便條舊格式
      else if (o && Array.isArray(o.items)) rows = o.items.map(function (n) { return { t: String(n.t || ''), d: !!n.d }; });
      else if (o && o.t) rows = String(o.t).split('\n').map(function (t) { return { t: t.trim(), d: false }; });           // 一整塊文字的舊格式
    } catch (e) {}
  }
  rows = rows.filter(function (r) { return r.t; });
  if (rows.length > BOARD_N) {                       // 超過 8 件，多的併進最後一格，不會不見
    var extra = rows.slice(BOARD_N - 1).map(function (r) { return r.t; }).join(' / ');
    rows = rows.slice(0, BOARD_N - 1).concat([{ t: extra, d: false }]);
  }
  while (rows.length < BOARD_N) rows.push({ t: '', d: false });
  return rows;
}
function adminBoard() { requireOwner_(); return getBoard_(); }
function adminBoardSave(items) {
  requireOwner_();
  if (!Array.isArray(items)) throw new Error('格式不對');
  var rows = items.slice(0, BOARD_N).map(function (n) { return { t: clean_(n && n.t, 300), d: !!(n && n.d) }; });
  while (rows.length < BOARD_N) rows.push({ t: '', d: false });
  boardSheet_().getRange('A1').setValue(JSON.stringify({ items: rows }));
  return rows;
}

function rememberWho_(cfg, who) {
  if (!who || REL_WORDS.indexOf(who) >= 0 || cfg.people.indexOf(who) >= 0) return false;
  cfg.people.unshift(who);
  cfg.people = cfg.people.slice(0, 200);
  return true;
}

function adminSaveAudience(aud, data) {
  requireOwner_();
  if (aud !== 'friend' && aud !== 'public') throw new Error('對象錯誤');
  var cfg = getConfig_();
  cfg.audiences[aud].minNoticeH = Math.max(0, Math.min(+data.minNoticeH || 0, 720));
  saveConfig_(cfg);
  return cfg;
}

/** 把同一組開放時段套用到多天；ranges 空陣列＝那幾天不開放 */
function adminSetOpen(dates, ranges) {
  requireOwner_();
  var cfg = getConfig_();
  var clean = (ranges || []).map(function (r) {
    if (!/^\d{1,2}:\d{2}$/.test(r.s) || !/^\d{1,2}:\d{2}$/.test(r.e) || toMin_(r.e) <= toMin_(r.s)) throw new Error('時間有誤：結束要晚於開始');
    return { s: r.s, e: r.e, loc: r.loc === 'online' ? 'online' : 'both', aud: ['friend', 'public'].indexOf(r.aud) >= 0 ? r.aud : 'all' };
  }).sort(function (a, b) { return toMin_(a.s) - toMin_(b.s); });
  (dates || []).forEach(function (d) {
    if (!/^\d{4}-\d{2}-\d{2}$/.test(d) || d > RANGE_END) throw new Error('日期錯誤');
    if (clean.length) cfg.open[d] = clean; else delete cfg.open[d];
  });
  var today = ymdOf_(new Date());
  Object.keys(cfg.open).forEach(function (k) { if (k < today) delete cfg.open[k]; });   // 清掉過去的
  saveConfig_(cfg);
  return cfg;
}

function adminCalendarSave(c) {
  requireOwner_();
  var cfg = getConfig_();
  var name = clean_(c.name, 40);
  if (!name) throw new Error('請填名稱');
  var color = /^#[0-9a-fA-F]{6}$/.test(c.color) ? c.color : '#4f7fd6';
  var kws = (c.keywords || []).map(function (k) { return clean_(k, 20); }).filter(String);
  if (c.id) {
    validCal_(cfg, c.id);
    var cal = CalendarApp.getCalendarById(c.id);
    cal.setName(name); cal.setColor(color);
    cfg.calendars.forEach(function (x) { if (x.id === c.id) { x.name = name; x.color = color; x.keywords = kws; } });
  } else {
    var nc = CalendarApp.createCalendar(name, { color: color, timeZone: TZ });
    cfg.calendars.push({ id: nc.getId(), name: name, color: color, keywords: kws });
  }
  saveConfig_(cfg);
  return cfg;
}

function adminCalendarMove(ids) {   // 調整順序（分類同分時，排前面的優先）
  requireOwner_();
  var cfg = getConfig_();
  var map = {}; cfg.calendars.forEach(function (c) { map[c.id] = c; });
  if (ids.length !== cfg.calendars.length || ids.some(function (id) { return !map[id]; })) throw new Error('順序有誤');
  cfg.calendars = ids.map(function (id) { return map[id]; });
  saveConfig_(cfg);
  return cfg;
}

function adminCalendarDelete(id, moveTo) {
  requireOwner_();
  var cfg = getConfig_();
  validCal_(cfg, id); validCal_(cfg, moveTo);
  if (id === moveTo) throw new Error('要移到另一個行事曆');
  var token = null, moved = 0, failed = 0;
  do {
    var opt = { singleEvents: false, maxResults: 2500, showDeleted: false };
    if (token) opt.pageToken = token;
    var res = Calendar.Events.list(id, opt);
    (res.items || []).forEach(function (ev) {
      if (ev.status === 'cancelled' || ev.recurringEventId) return;
      try { Calendar.Events.move(id, ev.id, moveTo); moved++; } catch (e) { failed++; }
    });
    token = res.nextPageToken;
  } while (token);
  if (failed) throw new Error('有 ' + failed + ' 個行程搬不過去，行事曆先保留。已搬 ' + moved + ' 個。');
  CalendarApp.getCalendarById(id).deleteCalendar();
  cfg.calendars = cfg.calendars.filter(function (c) { return c.id !== id; });
  ['friend', 'public'].forEach(function (a) { if (cfg.bookTo[a] === id) cfg.bookTo[a] = moveTo; });
  saveConfig_(cfg);
  return cfg;
}

function adminSaveShare(ids) {
  requireOwner_();
  var cfg = getConfig_();
  cfg.shareCals = (ids || []).filter(function (id) { return cfg.calendars.some(function (c) { return c.id === id; }); });
  cfg.shareVer = (cfg.shareVer || 0) + 1;     // 讓分享頁快取失效
  saveConfig_(cfg);
  return cfg;
}

/** 本週頁不顯示的行事曆（例如正職） */
function adminSaveWeekHide(ids) {
  requireOwner_();
  var cfg = getConfig_();
  cfg.weekHide = (ids || []).filter(function (id) { return cfg.calendars.some(function (c) { return c.id === id; }); });
  saveConfig_(cfg);
  return cfg.weekHide;
}

function adminSaveBookTo(bookTo) {
  requireOwner_();
  var cfg = getConfig_();
  validCal_(cfg, bookTo.friend); validCal_(cfg, bookTo.public);
  cfg.bookTo = { friend: bookTo.friend, public: bookTo.public };
  saveConfig_(cfg);
  return cfg;
}

/** 顯示在對外頁面的名字、品牌、一句話 */
function adminSaveMe(o) {
  requireOwner_();
  var cfg = getConfig_();
  cfg.me = { name: clean_(o.name, 20) || '我', brand: clean_(o.brand, 40), tagline: clean_(o.tagline, 60) };
  saveConfig_(cfg);
  return cfg;
}

function adminSaveLinks(o) {
  requireOwner_();
  var cfg = getConfig_();
  var url = clean_(o.publicUrl, 300);
  if (url && !/^https:\/\/script\.google\.com\/macros\/s\/[\w-]+\/exec$/.test(url)) throw new Error('對外網址格式不對，應該是 https://script.google.com/macros/s/…/exec');
  cfg.publicUrl = url;
  cfg.shortFriend = clean_(o.shortFriend, 200);
  cfg.shortPublic = clean_(o.shortPublic, 200);
  cfg.shortShare = clean_(o.shortShare, 200);
  if (o.notifyEmail && !isEmail_(o.notifyEmail)) throw new Error('通知信箱格式不對');
  cfg.notifyEmail = clean_(o.notifyEmail, 100) || cfg.notifyEmail;
  saveConfig_(cfg);
  return adminBoot();
}

function adminResetFriendToken() {
  requireOwner_();
  var cfg = getConfig_();
  cfg.audiences.friend.token = newToken_();
  cfg.shortFriend = '';        // 舊短網址會指向失效連結，一起清掉
  saveConfig_(cfg);
  return adminBoot();
}

function newToken_() { return Utilities.getUuid().replace(/-/g, '').slice(0, 12); }

function adminInbox() {
  requireOwner_();
  var sh = ss_().getSheetByName(INBOX);
  var n = sh.getLastRow();
  if (n < 2) return [];
  var start = Math.max(2, n - 499);
  var vals = sh.getRange(start, 1, n - start + 1, COLS.length).getDisplayValues();
  return vals.map(function (r) {
    return { id: r[0], at: r[1], status: r[2], type: r[3], aud: r[4], name: r[5], contact: r[6], when: r[7], dur: r[8],
      purpose: r[9], locMode: r[10], place: r[11], meet: r[12], calId: r[13], eventId: r[14] };
  }).reverse();
}

function findInboxRow_(id) {
  var sh = ss_().getSheetByName(INBOX);
  var ids = sh.getRange(1, 1, sh.getLastRow(), 1).getValues();
  for (var i = ids.length - 1; i >= 1; i--) if (String(ids[i][0]) === String(id)) return { sh: sh, row: i + 1 };
  throw new Error('找不到這筆');
}

function adminInboxStatus(id, status) {
  requireOwner_();
  if (['待處理', '待審核', '已排', '不約', '婉拒'].indexOf(status) < 0) throw new Error('狀態錯誤');
  var f = findInboxRow_(id);
  f.sh.getRange(f.row, 3).setValue(status);
  return adminInbox();
}

/** 通過待審核的預約：正式成立、發邀請、通知對方 */
function adminApprove(id) {
  requireOwner_();
  var cfg = getConfig_();
  var f = findInboxRow_(id);
  var r = f.sh.getRange(f.row, 1, 1, COLS.length).getDisplayValues()[0];
  if (r[2] !== '待審核') throw new Error('這筆不是待審核');
  var who = r[4], name = r[5], contact = r[6], when = r[7], dur = r[8], purpose = r[9];
  var online = r[10] === '線上', place = r[11], calId = r[13], evId = r[14];
  var locText = online ? '線上（Google Meet）' : place;
  var map = online ? '' : mapUrl_(place);
  var patch = { summary: '【' + who + '】' + name + '｜' + String(purpose).slice(0, 30),
    description: bookDesc_(who, name, contact, purpose, dur, locText, map, false) };
  var opt = {};
  if (online) { patch.conferenceData = { createRequest: { requestId: Utilities.getUuid(), conferenceSolutionKey: { type: 'hangoutsMeet' } } }; opt.conferenceDataVersion = 1; }
  if (isEmail_(contact)) { patch.attendees = [{ email: contact }]; opt.sendUpdates = 'all'; }
  var ev = Calendar.Events.patch(patch, calId, evId, opt);
  var meet = meetLink_(ev);
  f.sh.getRange(f.row, 3).setValue('已排');
  if (meet) f.sh.getRange(f.row, 13).setValue(meet);
  if (isEmail_(contact)) {
    sendCard_(contact, '預約成立｜' + when + '｜' + cfg.me.name, [['時間', when], ['多久', dur], ['約哪', locText], ['要幹嘛', purpose]],
      { tag: '已確認', heading: '約好了', when: when, from: cfg.me.name,
        buttons: [meet && ['加入 Google Meet', meet], map && ['在 Google 地圖打開', map]],
        foot: '要改時間或取消，直接回這封信說一聲。' });
  }
  return adminInbox();
}

/** 婉拒：放掉時段，並（可加一句話）通知對方 */
function adminReject(id, msg) {
  requireOwner_();
  var cfg = getConfig_();
  var f = findInboxRow_(id);
  var r = f.sh.getRange(f.row, 1, 1, COLS.length).getDisplayValues()[0];
  var contact = r[6], when = r[7], calId = r[13], evId = r[14];
  if (calId && evId) { try { Calendar.Events.remove(calId, evId); } catch (e) { console.error(e); } }
  f.sh.getRange(f.row, 3).setValue('婉拒');
  var note = clean_(msg, 500);
  if (isEmail_(contact)) {
    var link = bookLink_(cfg);
    sendCard_(contact, '這個時間沒辦法｜' + cfg.me.name, [['你原本約的', when]].concat(note ? [[cfg.me.name + ' 說', note]] : []),
      { tag: '很抱歉', heading: '這個時間不行', from: cfg.me.name,
        buttons: [link && ['挑別的時間', link]],
        foot: '不好意思，換個時間再約。' });
  }
  return adminInbox();
}

/** 存某個月要給預約者看的話：朋友和新朋友各一份 */
function adminSaveNote(ym, o) {
  requireOwner_();
  var cfg = getConfig_();
  if (!/^\d{4}-\d{2}$/.test(ym)) throw new Error('月份錯誤');
  if (!cfg.notes) cfg.notes = {};
  o = o || {};
  var f = note_(o.f).slice(0, 500), pb = note_(o.p).slice(0, 500);
  if (!f && !pb) delete cfg.notes[ym];
  else cfg.notes[ym] = { f: f, p: pb };
  saveConfig_(cfg);
  return cfg;
}

function adminSaveReview(aud, on) {
  requireOwner_();
  var cfg = getConfig_();
  if (['friend', 'public'].indexOf(aud) < 0) throw new Error('對象錯誤');
  if (!cfg.review) cfg.review = {};
  cfg.review[aud] = !!on;
  saveConfig_(cfg);
  return cfg;
}

function adminInboxSchedule(id, o) {
  requireOwner_();
  var cfg = getConfig_();
  var f = findInboxRow_(id);
  var r = f.sh.getRange(f.row, 1, 1, COLS.length).getDisplayValues()[0];
  var aud = r[4] === '朋友' ? 'friend' : 'public';
  var calId = o.cal || cfg.bookTo[aud];
  validCal_(cfg, calId);
  var s = +o.s, e = +o.e;
  if (!(e > s)) throw new Error('時間有誤');
  var online = o.locMode !== 'inperson';
  var locText = online ? '線上（Google Meet）' : clean_(o.place, 500);
  var contact = r[6];
  var ev = insertEvent_(calId, {
    title: '【' + r[4] + '】' + r[5] + '｜' + String(r[9]).slice(0, 30), ymd: o.date, s: s, e: e, loc: locText,
    desc: ['透過預約頁（' + r[4] + '，只約月份）', '誰：' + r[5], '聯絡：' + (contact || '—'), '要幹嘛：' + r[9], '約哪：' + locText].join('\n'),
    attendees: isEmail_(contact) ? [contact] : [], meet: online, who: r[5]
  });
  var when = whenText_(o.date, s, e);
  f.sh.getRange(f.row, 3).setValue('已排');
  f.sh.getRange(f.row, 8).setValue(r[7] + ' → ' + when);
  f.sh.getRange(f.row, 11, 1, 5).setValues([[online ? '線上' : '實體', online ? '' : locText, meetLink_(ev), calId, ev.id]]);
  return adminInbox();
}
