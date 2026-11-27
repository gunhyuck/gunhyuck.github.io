/* ============================================================
   합동 비서실 공용 스크립트
   - 초기 스냅샷으로 즉시 렌더 → Google Apps Script 연동 시 실시간 동기화
   - 연동 주소 우선순위: ?api=... > localStorage(설정 ⚙) > OFFICE_CONFIG.apiDefault
   ============================================================ */
(function () {
  "use strict";

  var CFG = window.OFFICE_CONFIG || {};
  var SITE = CFG.site || "karim";
  var STA = ["대기", "진행중", "검토중", "완료"];
  var ALL = "전체";

  var LS = {
    api: "office_api_url",
    cache: "office_cache_v1_" + SITE,
    author: "office_author_name",
    role: "office_author_role"
  };

  var state = {
    tasks: [],
    memos: [],
    status: ALL,
    q: "",
    category: ALL,
    api: "",
    live: false,
    busy: false,
    lastSync: null
  };

  /* ---------------- 유틸 ---------------- */
  function $(sel, root) { return (root || document).querySelector(sel); }
  function $$(sel, root) { return Array.prototype.slice.call((root || document).querySelectorAll(sel)); }
  function esc(s) {
    return String(s == null ? "" : s).replace(/[&<>"']/g, function (c) {
      return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c];
    });
  }
  function todayStr() {
    var d = new Date();
    var p = function (n) { return (n < 10 ? "0" : "") + n; };
    return d.getFullYear() + "-" + p(d.getMonth() + 1) + "-" + p(d.getDate());
  }
  function nowStamp() {
    var d = new Date();
    var p = function (n) { return (n < 10 ? "0" : "") + n; };
    return todayStr() + " " + p(d.getHours()) + ":" + p(d.getMinutes());
  }
  function parseDate(v) {
    if (!v) return null;
    var m = String(v).match(/(\d{4})[-.\/](\d{1,2})[-.\/](\d{1,2})/);
    if (!m) return null;
    return new Date(+m[1], +m[2] - 1, +m[3]);
  }
  function dueInfo(v) {
    var d = parseDate(v);
    if (!d) return { text: v ? String(v) : "-", cls: "" };
    var diff = Math.round((d - parseDate(todayStr())) / 86400000);
    var text = (d.getMonth() + 1) + "/" + d.getDate();
    if (diff < 0) return { text: text + " · 지남", cls: "past" };
    if (diff === 0) return { text: text + " · 오늘", cls: "soon" };
    if (diff <= 7) return { text: text + " · D-" + diff, cls: "soon" };
    return { text: text, cls: "" };
  }
  function fmtWhen(v) {
    var s = String(v || "").trim();
    var m = s.match(/^(\d{4})-(\d{2})-(\d{2})[ T]?(\d{2}:\d{2})?/);
    if (!m) return s;
    var d = new Date(+m[1], +m[2] - 1, +m[3]);
    var base = (d.getMonth() + 1) + "월 " + d.getDate() + "일";
    var today = parseDate(todayStr());
    var diff = Math.round((today - new Date(+m[1], +m[2] - 1, +m[3])) / 86400000);
    if (diff === 0) base = "오늘";
    else if (diff === 1) base = "어제";
    return base + (m[4] ? " " + m[4] : "");
  }
  function initials(name) {
    var n = String(name || "?").replace(/\(.*?\)/g, "").trim();
    var parts = n.split(/\s+/);
    if (parts.length > 1 && parts[parts.length - 1].length <= 2) return parts[0].slice(0, 1);
    return n.slice(0, 1) || "?";
  }

  /* ---------------- 설정값 ---------------- */
  function readApi() {
    var qs = new URLSearchParams(location.search);
    var fromQuery = qs.get("api");
    if (fromQuery) {
      try { localStorage.setItem(LS.api, fromQuery.trim()); } catch (e) {}
      return fromQuery.trim();
    }
    var stored = "";
    try { stored = localStorage.getItem(LS.api) || ""; } catch (e) {}
    return (stored || CFG.apiDefault || "").trim();
  }
  function saveApi(url) {
    try {
      if (url) localStorage.setItem(LS.api, url);
      else localStorage.removeItem(LS.api);
    } catch (e) {}
  }

  function readCache() {
    try {
      var raw = localStorage.getItem(LS.cache);
      if (!raw) return null;
      var o = JSON.parse(raw);
      if (o && Array.isArray(o.tasks) && Array.isArray(o.memos)) return o;
    } catch (e) {}
    return null;
  }
  function writeCache() {
    try {
      localStorage.setItem(LS.cache, JSON.stringify({
        tasks: state.tasks, memos: state.memos, ts: Date.now()
      }));
    } catch (e) {}
  }

  /* ---------------- 원격 연동 ---------------- */
  function withTimeout(promise, ms) {
    return new Promise(function (resolve, reject) {
      var t = setTimeout(function () { reject(new Error("timeout")); }, ms);
      promise.then(function (v) { clearTimeout(t); resolve(v); }, function (e) { clearTimeout(t); reject(e); });
    });
  }

  function jsonp(url, ms) {
    return new Promise(function (resolve, reject) {
      var cb = "officecb_" + Math.random().toString(36).slice(2);
      var s = document.createElement("script");
      var done = false;
      var t = setTimeout(function () { cleanup(); reject(new Error("timeout")); }, ms || 15000);
      function cleanup() {
        if (done) return;
        done = true;
        clearTimeout(t);
        try { delete window[cb]; } catch (e) { window[cb] = undefined; }
        if (s.parentNode) s.parentNode.removeChild(s);
      }
      window[cb] = function (data) { cleanup(); resolve(data); };
      s.onerror = function () { cleanup(); reject(new Error("script error")); };
      s.src = url + (url.indexOf("?") >= 0 ? "&" : "?") + "callback=" + cb;
      document.head.appendChild(s);
    });
  }

  async function fetchJson(url, opts) {
    var ctrl = new AbortController();
    var t = setTimeout(function () { ctrl.abort(); }, 12000);
    try {
      var r = await fetch(url, Object.assign({ cache: "no-store", redirect: "follow", signal: ctrl.signal }, opts || {}));
      if (!r.ok) throw new Error("HTTP " + r.status);
      var txt = await r.text();
      var data;
      try { data = JSON.parse(txt); } catch (e) { throw new Error("JSON 아님"); }
      return data;
    } finally { clearTimeout(t); }
  }

  function listUrl() {
    return state.api + (state.api.indexOf("?") >= 0 ? "&" : "?") + "action=list&site=" + encodeURIComponent(SITE) + "&_=" + Date.now();
  }

  async function remoteList() {
    var url = listUrl();
    try {
      var d = await fetchJson(url);
      if (d && d.ok !== false) return d;
    } catch (e) { /* CORS/네트워크 → JSONP 로 재시도 */ }
    return await withTimeout(jsonp(url, 15000), 16000);
  }

  async function remoteCall(payload) {
    // 새 안건·메모는 미리 만든 고유 id 를 함께 보냅니다.
    // POST 응답이 CORS 로 가려져 재전송하더라도 서버가 같은 id 를 중복 저장하지 않습니다.
    var body = JSON.stringify(payload);
    try {
      var d = await fetchJson(state.api, {
        method: "POST",
        headers: { "Content-Type": "text/plain;charset=utf-8" },
        body: body
      });
      if (d && typeof d === "object") return d;
    } catch (e) { /* CORS 차단 시 GET+JSONP 로 전송 */ }
    var url = state.api + (state.api.indexOf("?") >= 0 ? "&" : "?") +
      "action=" + encodeURIComponent(payload.action) +
      "&site=" + encodeURIComponent(payload.site || SITE) +
      "&data=" + encodeURIComponent(JSON.stringify(payload)) +
      "&_=" + Date.now();
    return await withTimeout(jsonp(url, 15000), 16000);
  }

  function applyRemote(d) {
    var tasks = (d.tasks || []).map(normalizeTask).filter(function (t) { return t.id || t.title; });
    var memos = (d.memos || []).map(normalizeMemo).filter(function (m) { return m.content; });
    state.tasks = tasks;
    state.memos = memos;
    state.live = true;
    state.lastSync = new Date();
    writeCache();
    renderAll();
  }

  function normalizeTask(t) {
    return {
      id: String(t.id || "").trim(),
      title: String(t.title || "").trim(),
      category: String(t.category || "").trim(),
      owner: String(t.owner || "").trim(),
      status: STA.indexOf(String(t.status || "").trim()) >= 0 ? String(t.status).trim() : "대기",
      due: String(t.due || "-").trim() || "-",
      summary: String(t.summary || "").trim(),
      updated_by: String(t.updated_by || "").trim(),
      updated_at: String(t.updated_at || "").trim()
    };
  }
  function normalizeMemo(m) {
    return {
      id: String(m.id || "").trim(),
      author: String(m.author || "익명").trim(),
      role: String(m.role || "").trim(),
      content: String(m.content || "").trim(),
      created_at: String(m.created_at || "").trim()
    };
  }

  /* ---------------- 화면 ---------------- */
  function statusPriority(s) {
    var i = ["진행중", "검토중", "대기", "완료"].indexOf(s);
    return i < 0 ? 9 : i;
  }
  function idNum(t) {
    var n = parseInt(String(t.id).replace(/[^0-9]/g, ""), 10);
    return isNaN(n) ? 9999 : n;
  }
  function sortedTasks(list) {
    return list.slice().sort(function (a, b) {
      return statusPriority(a.status) - statusPriority(b.status) || idNum(a) - idNum(b);
    });
  }
  function visibleTasks() {
    var q = state.q.trim().toLowerCase();
    return sortedTasks(state.tasks).filter(function (t) {
      if (state.status !== ALL && t.status !== state.status) return false;
      if (state.category !== ALL && t.category !== state.category) return false;
      if (!q) return true;
      return [t.id, t.title, t.category, t.owner, t.summary, t.status]
        .join(" ").toLowerCase().indexOf(q) >= 0;
    });
  }

  function renderStats() {
    var counts = {};
    counts[ALL] = state.tasks.length;
    STA.forEach(function (s) { counts[s] = 0; });
    state.tasks.forEach(function (t) { counts[t.status] = (counts[t.status] || 0) + 1; });

    $$(".stat").forEach(function (el) {
      var key = el.getAttribute("data-status");
      var v = el.querySelector(".v");
      var n = counts[key] || 0;
      var pct = state.tasks.length ? Math.round((n / state.tasks.length) * 100) : 0;
      v.innerHTML = n + (key === ALL ? "<small>건</small>" : "<small>" + pct + "%</small>");
      el.classList.toggle("on", state.status === key);
    });

    var ht = $("#heroTasks");
    if (ht) ht.textContent = state.tasks.length + "건";
  }

  function renderCategories() {
    var sel = $("#categoryFilter");
    if (!sel) return;
    var cats = [];
    state.tasks.forEach(function (t) {
      if (t.category && cats.indexOf(t.category) < 0) cats.push(t.category);
    });
    cats.sort();
    var html = '<option value="' + esc(ALL) + '">분류 전체</option>';
    cats.forEach(function (c) {
      html += '<option value="' + esc(c) + '">' + esc(c) + "</option>";
    });
    sel.innerHTML = html;
    if (cats.indexOf(state.category) < 0) state.category = ALL;
    sel.value = state.category;
  }

  function cardHtml(t) {
    var d = dueInfo(t.due);
    var cat = t.category ? '<span class="tag">' + esc(t.category) + "</span>" : "";
    var owner = t.owner
      ? '<div class="task-owner"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8"><circle cx="12" cy="8" r="3.4"/><path d="M4.8 20c.7-3.9 3.7-6 7.2-6s6.5 2.1 7.2 6"/></svg><span>' + esc(t.owner) + "</span></div>"
      : "";
    var opts = STA.map(function (s) {
      return '<option value="' + s + '"' + (s === t.status ? " selected" : "") + ">" + s + "</option>";
    }).join("");
    var footParts = [];
    if (t.due && t.due !== "-") {
      footParts.push('<span class="due ' + d.cls + '"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3.5" y="5" width="17" height="16" rx="3"/><path d="M8 3v4M16 3v4M3.5 10h17"/></svg>' + esc(d.text) + "</span>");
    } else {
      footParts.push('<span class="due">기한 미정</span>');
    }
    var who = [t.updated_by, t.updated_at ? fmtWhen(t.updated_at) : ""].filter(Boolean).join(" · ");

    return '' +
      '<article class="task-card' + (t.status === "완료" ? " is-done" : "") + '" data-id="' + esc(t.id) + '" data-status="' + esc(t.status) + '">' +
        '<div class="task-head">' +
          (t.id ? '<span class="tid">' + esc(t.id) + "</span>" : "") +
          cat +
          '<span class="spacer"></span>' +
          '<select class="status-select" data-v="' + esc(t.status) + '" data-id="' + esc(t.id) + '" aria-label="' + esc(t.title) + ' 상태 변경">' + opts + "</select>" +
        "</div>" +
        '<h3 class="task-title">' + esc(t.title || "(제목 없음)") + "</h3>" +
        owner +
        (t.summary ? '<p class="task-summary">' + esc(t.summary) + "</p>" : "") +
        '<div class="task-foot">' + footParts.join("") +
          '<span class="spacer"></span>' +
          (who ? "<span>" + esc(who) + "</span>" : "") +
        "</div>" +
      "</article>";
  }

  function renderTasks() {
    var grid = $("#taskGrid");
    if (!grid) return;
    var list = visibleTasks();
    if (!list.length) {
      var msg = state.tasks.length
        ? "조건에 맞는 안건이 없습니다."
        : "아직 등록된 안건이 없습니다. <b>+ 새 안건 추가</b>로 첫 안건을 남겨 주세요.";
      grid.innerHTML = '<div class="empty"><b>표시할 안건이 없습니다</b>' + msg + "</div>";
    } else {
      grid.innerHTML = list.map(cardHtml).join("");
    }
    var counter = $("#taskCount");
    if (counter) counter.textContent = list.length + " / " + state.tasks.length + "건";
  }

  function memoHtml(m) {
    return '' +
      '<div class="memo">' +
        '<div class="avatar">' + esc(initials(m.author)) + "</div>" +
        '<div class="memo-body">' +
          '<div class="memo-top">' +
            '<span class="memo-author">' + esc(m.author) + "</span>" +
            (m.role ? '<span class="memo-role">' + esc(m.role) + "</span>" : "") +
            '<span class="memo-time">' + esc(fmtWhen(m.created_at)) + "</span>" +
          "</div>" +
          '<div class="memo-text">' + esc(m.content) + "</div>" +
        "</div>" +
      "</div>";
  }

  function renderMemos() {
    var box = $("#memoList");
    if (!box) return;
    var list = state.memos.slice().sort(function (a, b) {
      return String(b.created_at).localeCompare(String(a.created_at));
    });
    box.innerHTML = list.length
      ? list.map(memoHtml).join("")
      : '<div class="memo" style="color:var(--ink-3);font-size:13px;">아직 메모가 없습니다. 첫 메모를 남겨 주세요.</div>';
    var c = $("#memoCount");
    if (c) c.textContent = list.length;
    var hm = $("#heroMemos");
    if (hm) hm.textContent = list.length + "건";
  }

  function renderSync() {
    var pill = $("#syncPill");
    if (!pill) return;
    pill.classList.remove("live", "local", "err");
    var txt = pill.querySelector(".txt");
    if (state.live) {
      pill.classList.add("live");
      txt.textContent = "구글 시트 연동";
      pill.title = "Google 시트와 연결됨 · 클릭하면 새로고침";
    } else if (state.busy) {
      txt.textContent = "연결 중…";
      pill.title = "구글 시트에 연결하는 중";
    } else if (state.api) {
      pill.classList.add("err");
      txt.textContent = "연동 확인 필요";
      pill.title = "연동에 실패해 저장된 내용을 표시 중입니다 · 클릭하면 다시 시도";
    } else {
      pill.classList.add("local");
      txt.textContent = "로컬 저장";
      pill.title = "이 브라우저에만 저장 중입니다 · ⚙에서 구글 시트 연동 주소를 넣어 주세요";
    }
    var st = $("#syncState");
    if (st) {
      st.textContent = state.live
        ? "연동됨" + (state.lastSync ? " · " + state.lastSync.toLocaleTimeString("ko-KR", { hour: "2-digit", minute: "2-digit" }) + " 동기화" : "")
        : "로컬 저장";
    }
  }

  function renderAll() {
    renderStats();
    renderCategories();
    renderTasks();
    renderMemos();
    renderSync();
  }

  /* ---------------- 토스트 ---------------- */
  function toast(msg, kind) {
    var stack = $("#toastStack");
    if (!stack) return;
    var el = document.createElement("div");
    el.className = "toast" + (kind ? " " + kind : "");
    el.textContent = msg;
    stack.appendChild(el);
    setTimeout(function () {
      el.style.transition = "opacity .25s, transform .25s";
      el.style.opacity = "0";
      el.style.transform = "translateY(6px)";
      setTimeout(function () { el.remove(); }, 260);
    }, 2600);
  }

  /* ---------------- 데이터 로드 ---------------- */
  async function refresh(manual) {
    if (!state.api) {
      state.live = false;
      renderSync();
      if (manual) toast("설정(⚙)에서 구글 시트 연동 주소를 먼저 넣어 주세요.", "err");
      return false;
    }
    state.busy = true;
    renderSync();
    try {
      var d = await remoteList();
      if (!d || d.ok === false) throw new Error((d && d.error) || "응답 오류");
      applyRemote(d);
      if (manual) toast("구글 시트와 동기화했습니다.", "ok");
      return true;
    } catch (e) {
      state.live = false;
      state.busy = false;
      renderSync();
      if (manual) toast("연동 실패: " + (e && e.message ? e.message : "알 수 없는 오류"), "err");
      return false;
    } finally {
      state.busy = false;
      renderSync();
    }
  }

  function bootstrap() {
    var snap = (window.OFFICE_DATA && window.OFFICE_DATA[SITE]) || { tasks: [], memos: [] };
    var cache = readCache();
    var source = cache && cache.tasks.length ? cache : snap;
    state.tasks = source.tasks.map(normalizeTask);
    state.memos = source.memos.map(normalizeMemo);
    state.api = readApi();
    state.live = false;
    renderAll();
    if (state.api) refresh(false);
  }

  /* ---------------- 작업 ---------------- */
  function myAuthor() {
    try { return localStorage.getItem(LS.author) || ""; } catch (e) { return ""; }
  }
  function setMyAuthor(name, role) {
    try {
      if (name) localStorage.setItem(LS.author, name);
      if (role) localStorage.setItem(LS.role, role);
    } catch (e) {}
  }

  async function changeStatus(id, next) {
    var t = state.tasks.filter(function (x) { return x.id === id; })[0];
    if (!t || t.status === next) return;
    var prev = t.status;
    var who = myAuthor() || "웹 비서실";
    t.status = next;
    t.updated_by = who;
    t.updated_at = nowStamp().slice(0, 10);
    writeCache();
    renderAll();
    if (!state.api) { toast("상태를 '" + next + "'(으)로 바꿨습니다. (이 브라우저에 저장)", "ok"); return; }
    try {
      var d = await remoteCall({ action: "updateStatus", site: SITE, id: id, status: next, updated_by: who });
      if (!d || d.ok === false) throw new Error((d && d.error) || "저장 실패");
      toast("'" + next + "'(으)로 저장했습니다.", "ok");
      refresh(false);
    } catch (e) {
      t.status = prev;
      writeCache();
      renderAll();
      toast("상태 저장 실패: " + (e && e.message ? e.message : "연결 오류"), "err");
    }
  }

  function nextId() {
    var nums = state.tasks.map(idNum).filter(function (n) { return n < 9999; });
    var n = nums.length ? Math.max.apply(null, nums) + 1 : 1;
    return (CFG.idPrefix || "T") + "-" + n;
  }

  async function addTask(data) {
    var who = data.updated_by || myAuthor() || "웹 비서실";
    var task = {
      id: nextId(),
      title: data.title,
      category: data.category,
      owner: data.owner,
      status: data.status,
      due: data.due || "-",
      summary: data.summary,
      updated_by: who,
      updated_at: todayStr()
    };
    setMyAuthor(who, "");
    if (!state.api) {
      state.tasks.push(task);
      writeCache();
      renderAll();
      toast("새 안건을 추가했습니다. (이 브라우저에 저장)", "ok");
      return;
    }
    try {
      var d = await remoteCall({ action: "addTask", site: SITE, task: task });
      if (!d || d.ok === false) throw new Error((d && d.error) || "저장 실패");
      toast("새 안건을 구글 시트에 추가했습니다." + (d.id ? " (" + d.id + ")" : ""), "ok");
      await refresh(false);
    } catch (e) {
      task.id = nextId();
      state.tasks.push(task);
      writeCache();
      renderAll();
      toast("시트 저장 실패로 이 브라우저에만 남겼습니다: " + (e && e.message ? e.message : "연결 오류"), "err");
    }
  }

  async function addMemo(data) {
    var memo = {
      id: "M-" + (CFG.idPrefix || "T") + "-" + Date.now().toString(36) + Math.random().toString(36).slice(2, 6),
      author: data.author,
      role: data.role,
      content: data.content,
      created_at: nowStamp()
    };
    setMyAuthor(memo.author, memo.role);
    if (!state.api) {
      state.memos.push(memo);
      writeCache();
      renderMemos();
      toast("메모를 남겼습니다. (이 브라우저에 저장)", "ok");
      return;
    }
    try {
      var d = await remoteCall({ action: "addMemo", site: SITE, memo: memo });
      if (!d || d.ok === false) throw new Error((d && d.error) || "저장 실패");
      toast("메모를 구글 시트에 남겼습니다.", "ok");
      await refresh(false);
    } catch (e) {
      state.memos.push(memo);
      writeCache();
      renderMemos();
      toast("시트 저장 실패로 이 브라우저에만 남겼습니다.", "err");
    }
  }

  /* ---------------- 모달 ---------------- */
  function openModal(id) {
    var m = document.getElementById(id);
    if (!m) return;
    m.classList.add("open");
    document.body.style.overflow = "hidden";
    var first = m.querySelector("input:not([type=hidden]), textarea, select");
    if (first) setTimeout(function () { first.focus(); }, 60);
  }
  function closeModal(m) {
    var el = typeof m === "string" ? document.getElementById(m) : m;
    if (!el) return;
    el.classList.remove("open");
    document.body.style.overflow = "";
  }

  /* ---------------- 이벤트 ---------------- */
  function bind() {
    $$(".stat").forEach(function (el) {
      el.addEventListener("click", function () {
        state.status = el.getAttribute("data-status");
        renderStats();
        renderTasks();
        var grid = $("#taskGrid");
        if (grid) grid.scrollIntoView({ behavior: "smooth", block: "nearest" });
      });
    });

    var q = $("#searchInput");
    if (q) q.addEventListener("input", function () { state.q = q.value; renderTasks(); });

    var cf = $("#categoryFilter");
    if (cf) cf.addEventListener("change", function () { state.category = cf.value; renderTasks(); });

    var grid = $("#taskGrid");
    if (grid) grid.addEventListener("change", function (e) {
      var sel = e.target.closest(".status-select");
      if (sel) changeStatus(sel.getAttribute("data-id"), sel.value);
    });

    var pill = $("#syncPill");
    if (pill) pill.addEventListener("click", function () { refresh(true); });

    var gear = $("#gearBtn");
    if (gear) gear.addEventListener("click", function () {
      var input = $("#apiInput");
      if (input) input.value = state.api || "";
      var st = $("#settingsState");
      if (st) st.textContent = state.live ? "현재 구글 시트와 연동되어 있습니다." : (state.api ? "연동 주소는 저장되어 있지만 아직 확인되지 않았습니다." : "아직 연동 주소가 없습니다. 두 페이지 모두 로컬 저장으로 동작합니다.");
      openModal("settingsModal");
    });

    $$("[data-close]").forEach(function (el) {
      el.addEventListener("click", function () { closeModal(el.getAttribute("data-close")); });
    });
    $$(".modal").forEach(function (m) {
      m.addEventListener("click", function (e) { if (e.target === m) closeModal(m); });
    });

    var addBtn = $("#addTaskBtn");
    if (addBtn) addBtn.addEventListener("click", function () {
      var f = $("#taskForm");
      f.reset();
      $("#taskDue").value = "";
      $("#taskAuthor").value = myAuthor() || CFG.defaultAuthor || "";
      openModal("taskModal");
    });

    var tf = $("#taskForm");
    if (tf) tf.addEventListener("submit", async function (e) {
      e.preventDefault();
      var title = $("#taskTitle").value.trim();
      if (!title) { toast("안건 제목을 입력해 주세요.", "err"); $("#taskTitle").focus(); return; }
      var btn = $("#taskSubmit");
      btn.disabled = true;
      await addTask({
        title: title,
        category: $("#taskCategory").value.trim(),
        owner: $("#taskOwner").value.trim(),
        status: $("#taskStatus").value,
        due: $("#taskDue").value.trim(),
        summary: $("#taskSummary").value.trim(),
        updated_by: $("#taskAuthor").value.trim()
      });
      btn.disabled = false;
      closeModal("taskModal");
    });

    var mf = $("#memoForm");
    if (mf) {
      var sel = $("#memoWho");
      var custom = $("#memoWhoCustom");
      var customWrap = $("#memoWhoCustomWrap");
      sel.addEventListener("change", function () {
        var isCustom = sel.value === "__custom";
        if (customWrap) customWrap.style.display = isCustom ? "" : "none";
        if (isCustom && custom) custom.focus();
      });
      mf.addEventListener("submit", async function (e) {
        e.preventDefault();
        var name = sel.value === "__custom" ? (custom ? custom.value.trim() : "") : sel.value;
        var content = $("#memoText").value.trim();
        if (!name) { toast("이름을 입력해 주세요.", "err"); return; }
        if (!content) { toast("메모 내용을 입력해 주세요.", "err"); $("#memoText").focus(); return; }
        var btn = $("#memoSubmit");
        btn.disabled = true;
        await addMemo({ author: name, role: $("#memoRole").value.trim(), content: content });
        btn.disabled = false;
        $("#memoText").value = "";
      });
    }

    var save = $("#saveApi");
    if (save) save.addEventListener("click", async function () {
      var v = $("#apiInput").value.trim();
      if (v && !/^https:\/\/script\.google\.com\/macros\/s\/[^/]+\/exec$/.test(v)) {
        if (!/^https:\/\/script\.google\.com\//.test(v)) {
          toast("Apps Script 웹 앱 주소(https://script.google.com/macros/s/.../exec)를 넣어 주세요.", "err");
          return;
        }
      }
      saveApi(v);
      state.api = v;
      if (!v) {
        state.live = false;
        state.lastSync = null;
        renderSync();
        toast("연동을 해제했습니다. 이제 이 브라우저에만 저장됩니다.", "ok");
        return;
      }
      save.disabled = true;
      save.textContent = "연결 확인 중…";
      var ok = await refresh(false);
      save.disabled = false;
      save.textContent = "저장하고 연결";
      if (ok) { toast("구글 시트와 연결되었습니다.", "ok"); closeModal("settingsModal"); }
      else toast("주소는 저장했지만 연결되지 않았습니다. 배포 권한이 '모든 사용자'인지 확인해 주세요.", "err");
    });

    var clr = $("#clearApi");
    if (clr) clr.addEventListener("click", function () {
      saveApi("");
      state.api = "";
      state.live = false;
      renderSync();
      $("#apiInput").value = "";
      toast("연동 주소를 지웠습니다.", "ok");
    });

    document.addEventListener("keydown", function (e) {
      if (e.key === "Escape") $$(".modal.open").forEach(closeModal);
    });
  }

  /* ---------------- 시작 ---------------- */
  document.addEventListener("DOMContentLoaded", function () {
    bind();
    bootstrap();
    setInterval(function () { if (state.api && !state.busy) refresh(false); }, 60000);
  });
})();
