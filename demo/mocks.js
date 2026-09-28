/* 在瀏覽器裡模擬 Google 的服務，讓 Code.gs 原封不動跑起來 */
function makeGoogle(owner) {
  var st = { props: {}, cache: {}, sheets: {}, cals: {}, events: {}, mails: [], seq: 0 };
  function sheet(name) {
    var rows = [], s = { rows: rows };
    s.getLastRow = function () { return rows.length; };
    s.appendRow = function (r) { rows.push(r.slice()); };
    s.setFrozenRows = s.setColumnWidth = s.hideSheet = function () {};
    s.deleteRow = function (i) { rows.splice(i - 1, 1); };
    s.setName = function (n) { delete st.sheets[name]; name = n; st.sheets[n] = s; return s; };
    s.getRange = function (a, b, c, d) {
      if (typeof a === 'string') return { getValue: function () { return (rows[0] || [])[0] || ''; }, setValue: function (v) { rows[0] = [v]; } };
      var r = a, col = b, nr = c || 1, nc = d || 1;
      var g = {
        setFontWeight: function () { return g; }, setBackground: function () { return g; }, setFontColor: function () { return g; },
        getValues: function () { var o = []; for (var i = 0; i < nr; i++) { o.push([]); for (var j = 0; j < nc; j++) { var v = (rows[r - 1 + i] || [])[col - 1 + j]; o[i].push(v == null ? '' : v); } } return o; },
        getDisplayValues: function () { return g.getValues().map(function (x) { return x.map(function (v) { return String(v).replace(/^'/, ''); }); }); },
        setValue: function (v) { rows[r - 1][col - 1] = v; },
        setValues: function (vs) { vs.forEach(function (row, i) { row.forEach(function (v, j) { rows[r - 1 + i][col - 1 + j] = v; }); }); }
      };
      return g;
    };
    st.sheets[name] = s; return s;
  }
  var ss = { getSheetByName: function (n) { return st.sheets[n] || null; }, getSheets: function () { return Object.keys(st.sheets).map(function (k) { return st.sheets[k]; }); },
    insertSheet: function (n) { return sheet(n); }, getUrl: function () { return 'https://docs.google.com/spreadsheets/d/DEMO'; }, getId: function () { return 'DEMO'; } };
  sheet('工作表1');

  function parts(d) {   // 試玩版一律用裝置當地時間，避免時區不同造成錯位
    var z = function (n) { return (n < 10 ? '0' : '') + n; };
    return { year: String(d.getFullYear()), month: z(d.getMonth() + 1), day: z(d.getDate()), hour: z(d.getHours()), minute: z(d.getMinutes()) };
  }
  var Utilities = {
    formatDate: function (d, tz, f) {
      var p = parts(d);
      if (f === 'yyyy-MM-dd') return p.year + '-' + p.month + '-' + p.day;
      if (f === 'H') return String(+p.hour); if (f === 'm') return String(+p.minute);
      return p.year + '-' + p.month + '-' + p.day + ' ' + p.hour + ':' + p.minute;
    },
    getUuid: function () { return 'xxxxxxxx-xxxx-4xxx'.replace(/x/g, function () { return (Math.random() * 16 | 0).toString(16); }); }
  };

  /* ---- 行事曆事件（含每週重複展開） ---- */
  var BY = { SU: 0, MO: 1, TU: 2, WE: 3, TH: 4, FR: 5, SA: 6 };
  function ms(x) { return x.dateTime ? new Date(x.dateTime.slice(0, 19)).getTime() : new Date(x.date + 'T00:00:00').getTime(); }
  function pad(n) { return (n < 10 ? '0' : '') + n; }
  function ymdL(d) { return d.getFullYear() + '-' + pad(d.getMonth() + 1) + '-' + pad(d.getDate()); }
  function shiftTo(x, day) { return x.dateTime ? { dateTime: day + x.dateTime.slice(10), timeZone: x.timeZone } : { date: day }; }
  function expand(ev, from, to) {
    if (!ev.recurrence) return [ev];
    var exd = [];
    ev.recurrence.forEach(function (x) { var mm = /^EXDATE[^:]*:(.*)$/.exec(x); if (mm) mm[1].split(',').forEach(function (d) { exd.push(d.slice(0, 8)); }); });
    var rule = ev.recurrence[0], days = /BYDAY=([A-Z,]+)/.exec(rule)[1].split(',').map(function (k) { return BY[k]; });
    var until = /UNTIL=(\d{4})(\d{2})(\d{2})/.exec(rule), untilMs = until ? new Date(until[1] + '-' + until[2] + '-' + until[3] + 'T23:59:59').getTime() : Infinity;
    var s0 = ms(ev.start), len = ms(ev.end) - s0, out = [];
    var d = new Date(s0); d.setHours(0, 0, 0, 0);
    var start = Math.max(d.getTime(), from - 86400000 * 2);
    for (var t = new Date(start); t.getTime() <= Math.min(to, untilMs); t.setDate(t.getDate() + 1)) {
      if (t.getTime() < d.getTime() || days.indexOf(t.getDay()) < 0) continue;
      var day = ymdL(t), iid = ev.id + '_' + day.replace(/-/g, '');
      if ((ev.exdates || []).indexOf(iid) >= 0 || exd.indexOf(day.replace(/-/g, '')) >= 0) continue;
      var st0 = shiftTo(ev.start, day), sMs = ms(st0);
      if (sMs + len <= from || sMs >= to) continue;
      var en = ev.start.dateTime ? { dateTime: '', timeZone: ev.end.timeZone } : { date: '' };
      var e = new Date(sMs + len);
      if (en.dateTime !== undefined) en.dateTime = ymdL(e) + 'T' + pad(e.getHours()) + ':' + pad(e.getMinutes()) + ':00'; else en.date = ymdL(e);
      out.push(Object.assign({}, ev, { id: iid, recurringEventId: ev.id, recurrence: undefined, start: st0, end: en }));
    }
    return out;
  }
  function all(cal, from, to) {
    var out = [];
    (st.events[cal] || []).forEach(function (ev) {
      expand(ev, from, to).forEach(function (x) { if (ms(x.end) > from && ms(x.start) < to) out.push(x); });
    });
    return out.sort(function (a, b) { return ms(a.start) - ms(b.start); });
  }
  function findCal(id) { for (var c in st.events) if (st.events[c].some(function (e) { return e.id === id; })) return c; return null; }
  var Calendar = {
    CalendarList: { list: function () {
      var items = Object.keys(st.cals).map(function (id) { return { id: id, summary: st.cals[id].name, accessRole: 'owner' }; });
      return { items: items.concat([{ id: 'family-demo@group.calendar.google.com', summary: '家庭', accessRole: 'writer' }, { id: owner, summary: owner, accessRole: 'owner', primary: true }]) };
    } },
    Freebusy: { query: function (q) {
      var from = new Date(q.timeMin).getTime(), to = new Date(q.timeMax).getTime(), res = {};
      q.items.forEach(function (i) {
        res[i.id] = { busy: all(i.id, from, to).filter(function (e) { return e.start.dateTime || e.transparency === 'opaque'; })
          .map(function (e) { return { start: new Date(ms(e.start)).toISOString(), end: new Date(ms(e.end)).toISOString() }; }) };
      });
      return { calendars: res };
    } },
    Events: {
      insert: function (ev, cal, opt) {
        var e = JSON.parse(JSON.stringify(ev)); e.id = 'ev' + (++st.seq); e.status = 'confirmed';
        if (opt && opt.conferenceDataVersion) e.hangoutLink = 'https://meet.google.com/' + Math.random().toString(36).slice(2, 5) + '-' + Math.random().toString(36).slice(2, 6) + '-' + Math.random().toString(36).slice(2, 5);
        delete e.conferenceData; e.htmlLink = 'https://calendar.google.com/calendar/event?eid=' + e.id;
        (st.events[cal] = st.events[cal] || []).push(e); return e;
      },
      list: function (cal, o) {
        if (!o.singleEvents) return { items: (st.events[cal] || []).slice() };
        return { items: all(cal, new Date(o.timeMin).getTime(), new Date(o.timeMax).getTime()) };
      },
      patch: function (p, cal, id) {
        var list = st.events[cal] || [], e = list.filter(function (x) { return x.id === id; })[0];
        if (!e) {       // 重複行程的某一次：拆成獨立例外
          var mid = id.split('_')[0], m = list.filter(function (x) { return x.id === mid; })[0];
          if (!m) throw new Error('找不到行程');
          (m.exdates = m.exdates || []).push(id);
          e = JSON.parse(JSON.stringify(m)); delete e.recurrence; delete e.exdates; e.id = 'ev' + (++st.seq); list.push(e);
        }
        ['summary', 'location', 'description', 'extendedProperties'].forEach(function (k) { if (p[k] !== undefined) e[k] = p[k]; });
        if (p.conferenceData && !e.hangoutLink) e.hangoutLink = 'https://meet.google.com/' + Math.random().toString(36).slice(2, 5) + '-' + Math.random().toString(36).slice(2, 6) + '-' + Math.random().toString(36).slice(2, 5);
        if (p.start) e.start = p.start.date ? { date: p.start.date } : { dateTime: p.start.dateTime, timeZone: p.start.timeZone };
        if (p.end) e.end = p.end.date ? { date: p.end.date } : { dateTime: p.end.dateTime, timeZone: p.end.timeZone };
        return e;
      },
      move: function (cal, id, dest) {
        var list = st.events[cal] || [], i = -1;
        list.forEach(function (x, k) { if (x.id === id) i = k; });
        if (i < 0) { var c = findCal(id); if (!c) throw new Error('找不到行程'); list = st.events[c]; list.forEach(function (x, k) { if (x.id === id) i = k; }); }
        var e = list.splice(i, 1)[0]; (st.events[dest] = st.events[dest] || []).push(e);
      },
      get: function (cal, id) {
        var e = (st.events[cal] || []).filter(function (x) { return x.id === id; })[0];
        if (!e) throw new Error('Not Found'); return e;
      },
      remove: function (cal, id) {
        var list = st.events[cal] || [], before = list.length;
        st.events[cal] = list.filter(function (x) { return x.id !== id; });
        if (st.events[cal].length === before) {
          var mid = id.split('_')[0], m = st.events[cal].filter(function (x) { return x.id === mid; })[0];
          if (m) (m.exdates = m.exdates || []).push(id);
        }
      }
    }
  };
  var CalendarApp = {
    getCalendarsByName: function () { return []; },
    createCalendar: function (n, o) { var id = 'cal' + (++st.seq) + '@group.calendar.google.com'; st.cals[id] = { name: n, color: o.color }; st.events[id] = []; return { getId: function () { return id; } }; },
    getCalendarById: function (id) { return { setName: function (n) { st.cals[id].name = n; }, setColor: function (c) { st.cals[id].color = c; }, deleteCalendar: function () { delete st.cals[id]; delete st.events[id]; } }; }
  };
  st.events[owner] = [];
  return {
    st: st,
    PropertiesService: { getScriptProperties: function () { return { getProperty: function (k) { return st.props[k] || null; }, setProperty: function (k, v) { st.props[k] = v; } }; } },
    CacheService: { getScriptCache: function () { return { get: function (k) { return st.cache[k] || null; }, put: function (k, v) { st.cache[k] = v; }, remove: function (k) { delete st.cache[k]; } }; } },
    SpreadsheetApp: { create: function () { return ss; }, openById: function () { return ss; } },
    Session: { getActiveUser: function () { return { getEmail: function () { return owner; } }; }, getEffectiveUser: function () { return { getEmail: function () { return owner; } }; } },
    LockService: { getScriptLock: function () { return { tryLock: function () { return true; }, releaseLock: function () {} }; } },
    MailApp: { sendEmail: function (o) { st.mails.unshift({ to: o.to, subject: o.subject, body: o.body, html: o.htmlBody, at: new Date() }); } },
    Logger: { log: function () {} },
    Utilities: Utilities, Calendar: Calendar, CalendarApp: CalendarApp,
    ScriptApp: {
      getService: function () { return { getUrl: function () { return 'https://script.google.com/macros/s/DEMO/exec'; } }; },
      getProjectTriggers: function () { return (st.triggers || []).map(function (h) { return { getHandlerFunction: function () { return h; } }; }); },
      deleteTrigger: function (t) { st.triggers = (st.triggers || []).filter(function (h) { return h !== t.getHandlerFunction(); }); },
      newTrigger: function (h) { var b = { timeBased: function () { return b; }, everyDays: function () { return b; }, atHour: function () { return b; }, inTimezone: function () { return b; }, create: function () { (st.triggers = st.triggers || []).push(h); } }; return b; }
    },
    HtmlService: {}
  };
}
