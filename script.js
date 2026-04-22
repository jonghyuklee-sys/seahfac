var isRefresh = sessionStorage.getItem('seah_session_active');
sessionStorage.setItem('seah_session_active', 'true');

var curLine = isRefresh ? (localStorage.getItem('seah_curLine') || 'DASH') : 'DASH';
var curPart = isRefresh ? (localStorage.getItem('seah_curPart') || 'mechanical') : 'mechanical';
var curLoc = isRefresh ? (localStorage.getItem('seah_curLoc') || 'ALL') : 'ALL';
var curFreq = isRefresh ? (localStorage.getItem('seah_curFreq') || 'ALL') : 'ALL';
var curDate = '';
var editing = false;
var curPeriod = isRefresh ? (localStorage.getItem('seah_curPeriod') || 'D') : 'D';
var curObsLine = isRefresh ? (localStorage.getItem('seah_curObsLine') || 'ALL') : 'ALL';
var curObsPage = 1;
var curMainPage = 1;
var saved = JSON.parse(localStorage.getItem('seah_insp') || '{}');
var customItems = JSON.parse(localStorage.getItem('seah_custom') || '{}');
var metaOverrides = JSON.parse(localStorage.getItem('seah_meta') || '{}');
var savedBackup = null, customBackup = null, metaBackup = null; // To revert if not saved

// Migration: clear all electrical metaOverrides (old placeholder data replaced by real Excel import)
(function() {
  if (localStorage.getItem('seah_elec_v') === '3') return;
  var lines = ['CPL','ARP','CRM','CGL','1CCL','2CCL','3CCL','SSCL'];
  var changed = false;
  lines.forEach(function(ln) {
    var ckey = ln + '_electrical';
    if (metaOverrides[ckey]) {
      delete metaOverrides[ckey];
      changed = true;
    }
  });
  if (changed) {
    localStorage.setItem('seah_meta', JSON.stringify(metaOverrides));
    window._elecMigrationPending = true;
  }
  localStorage.setItem('seah_elec_v', '3');
})();

var curPhotoKey = null;
var curPhotoIdx = null;
var curPhotoDay = null;
var curViewerIdx = 0;

function getWeekOfMonth(dateStr) {
  var d = new Date(dateStr + 'T00:00:00');
  var date = d.getDate();
  return Math.floor((date - 1) / 7) + 1;
}

function getBaseItems(line, part) {
  if (part === 'mechanical') return (DATA[line] || []);
  return (DATA_ELEC[line] || []);
}

// Global Keyboard listener for ESC
document.addEventListener('keydown', function (e) {
  if (e.key === 'Escape') {
    closePhotoModal();
    if (document.getElementById('addModal')) document.getElementById('addModal').classList.remove('show');
    closePwdModal();
  }
});

// Firebase Configuration
const firebaseConfig = {
  apiKey: "AIzaSyD7IwzEihBPm3nYDersTjeDRMONmmyqh98",
  authDomain: "facility-common.firebaseapp.com",
  projectId: "facility-common",
  storageBucket: "facility-common.firebasestorage.app",
  messagingSenderId: "28758728108",
  appId: "1:28758728108:web:1a648edb9fd42d90b56e09"
};
firebase.initializeApp(firebaseConfig);
const db = firebase.firestore();
const storage = firebase.storage();

if (window._elecMigrationPending) {
  db.collection('settings').doc('metaOverrides').set(metaOverrides).catch(function(){});
  window._elecMigrationPending = false;
}

function init() {
  var t = new Date();
  var today = t.getFullYear() + '-' + String(t.getMonth() + 1).padStart(2, '0') + '-' + String(t.getDate()).padStart(2, '0');
  if (!curDate) curDate = today;
  document.getElementById('datePicker').value = curDate;
  window.curSearch = '';

  document.getElementById('lineNav').addEventListener('click', function (e) {
    var btn = e.target;
    while (btn && !btn.classList.contains('nav-btn')) btn = btn.parentElement;
    if (!btn) return;
    var ln = btn.getAttribute('data-line');
    if (ln === 'PREDICT') {
      alert('준비 중 입니다. 데이터 축적 후 서비스 예정입니다.');
      return;
    }
    goLine(ln);
  });

  document.getElementById('datePicker').addEventListener('change', function (e) {
    curDate = e.target.value;
    localStorage.setItem('seah_curDate', curDate);
    curObsPage = 1;
    curMainPage = 1;
    fetchWeekData(); // Fetch new week data if date changes
    render();
  });

  document.getElementById('editBtn').addEventListener('click', function () {
    if (!editing) {
      // 세션에 인증 정보가 있는 경우 바로 수정 모드 진입
      if (sessionStorage.getItem('seah_authenticated') === 'true') {
        savedBackup = JSON.stringify(saved);
        customBackup = JSON.stringify(customItems);
        metaBackup = JSON.stringify(metaOverrides);
        editing = true;
        updateEditUI(); render();
      } else {
        document.getElementById('pwdModal').classList.add('show');
        document.getElementById('editPwdInp').value = '';
        setTimeout(() => document.getElementById('editPwdInp').focus(), 100);
      }
    } else {
      if (savedBackup !== null) {
        saved = JSON.parse(savedBackup);
        customItems = JSON.parse(customBackup);
        metaOverrides = JSON.parse(metaBackup);
        savedBackup = null; customBackup = null; metaBackup = null;
      }
      editing = false;
      updateEditUI(); render();
    }
  });

  // Handle Enter key in password modal
  document.getElementById('editPwdInp').addEventListener('keypress', function(e) {
    if (e.key === 'Enter') confirmPwd();
  });
  document.getElementById('saveBtn').addEventListener('click', saveData);
  document.getElementById('exportBtn').addEventListener('click', exportExcel);
  document.getElementById('importBtn').addEventListener('click', function () {
    document.getElementById('excelUpload').click();
  });
  document.getElementById('excelUpload').addEventListener('change', importExcel);

  document.getElementById('addRowBtn').addEventListener('click', function () {
    document.getElementById('addModal').classList.add('show');
    document.getElementById('newLoc').value = curLoc !== 'ALL' ? curLoc : '';
    document.getElementById('newEquip').value = '';
    document.getElementById('newFreq').value = 'D';
    document.getElementById('newLoc').focus();
  });
  document.getElementById('cancelAddBtn').addEventListener('click', function () {
    document.getElementById('addModal').classList.remove('show');
  });
  document.getElementById('confirmAddBtn').addEventListener('click', addItem);

  document.getElementById('obsAddBtn').addEventListener('click', addObsRow);

  document.getElementById('tbody').addEventListener('click', function (e) {
    if (e.target.classList.contains('btn-del')) {
      var idx = parseInt(e.target.getAttribute('data-idx'));
      deleteItem(idx);
    }
  });

  setupDropdowns();
  updatePartBtnLabel();
  buildPartDropdown();
  buildLocDropdown();
  buildFreqDropdown();
  buildObsLineDropdown();

  for (var k in saved) { if (k.startsWith('OBS_')) delete saved[k]; }
  localStorage.setItem('seah_insp', JSON.stringify(saved));

  goLine(curLine, true);

  // 자동 한 번 일괄 클라우드 업로드 (처음 1회만)
  if (!localStorage.getItem('seah_migrated')) {
    const keys = Object.keys(saved);
    if (keys.length > 0) {
      console.log('클라우드 자동 이식 시작...');
      Promise.all([
        ...keys.map(k => db.collection('inspections').doc(k).set(saved[k])),
        db.collection('settings').doc('customItems').set(customItems)
      ]).then(() => {
        localStorage.setItem('seah_migrated', 'done');
        console.log('클라우드 이식 완료!');
      }).catch(err => console.error("Auto Sync Error:", err));
    }
  }
  // Mobile Menu Toggle
  const menuToggle = document.getElementById('menuToggle');
  const sidebar = document.getElementById('sidebar');
  if (menuToggle && sidebar) {
    menuToggle.addEventListener('click', function(e) {
      e.stopPropagation();
      sidebar.classList.toggle('opened');
      menuToggle.classList.toggle('active');
    });
    
    // Close menu when clicking outside
    document.addEventListener('click', function(e) {
      if (sidebar.classList.contains('opened') && !sidebar.contains(e.target)) {
        sidebar.classList.remove('opened');
        menuToggle.classList.remove('active');
      }
    });
  }
}

function setupDropdowns() {
  document.addEventListener('click', function (e) {
    var dds = document.querySelectorAll('.dropdown');
    for (var i = 0; i < dds.length; i++) {
      if (!dds[i].contains(e.target)) dds[i].classList.remove('open');
    }
  });
  document.getElementById('btnPart').addEventListener('click', function (e) {
    e.stopPropagation();
    closeOtherDropdowns('ddPart');
    document.getElementById('ddPart').classList.toggle('open');
  });
  document.getElementById('btnLoc').addEventListener('click', function (e) {
    e.stopPropagation();
    closeOtherDropdowns('ddLoc');
    document.getElementById('ddLoc').classList.toggle('open');
  });
  document.getElementById('btnFreq').addEventListener('click', function (e) {
    e.stopPropagation();
    closeOtherDropdowns('ddFreq');
    document.getElementById('ddFreq').classList.toggle('open');
  });
  if (document.getElementById('btnObsLine')) {
    document.getElementById('btnObsLine').addEventListener('click', function (e) {
      e.stopPropagation();
      closeOtherDropdowns('ddObsLine');
      document.getElementById('ddObsLine').classList.toggle('open');
    });
  }
}

function closeOtherDropdowns(except) {
  var ids = ['ddPart', 'ddLoc', 'ddFreq', 'ddObsLine'];
  for (var i = 0; i < ids.length; i++) {
    var el = document.getElementById(ids[i]);
    if (el && ids[i] !== except) el.classList.remove('open');
  }
}

function buildPartDropdown() {
  var parts = [{ v: 'mechanical', l: '기계' }, { v: 'electrical', l: '전기' }];
  var h = '';
  for (var i = 0; i < parts.length; i++) {
    h += '<div class="dropdown-item' + (parts[i].v === curPart ? ' active' : '') + '" data-val="' + parts[i].v + '">' + parts[i].l + '</div>';
  }
  var list = document.getElementById('listPart');
  list.innerHTML = h;
  list.addEventListener('click', function (e) {
    var item = e.target.closest('.dropdown-item');
    if (!item) return;
    setPart(item.getAttribute('data-val'));
  });
}

function setPart(p) {
  curPart = p;
  localStorage.setItem('seah_curPart', p);
  updatePartBtnLabel();
  document.getElementById('ddPart').classList.remove('open');
  curLoc = 'ALL'; localStorage.setItem('seah_curLoc', 'ALL');
  curFreq = 'ALL'; localStorage.setItem('seah_curFreq', 'ALL');
  curMainPage = 1;
  editing = false; updateEditUI(); buildLocDropdown(); buildFreqDropdown(); render();
}

function updatePartBtnLabel() {
  var label = curPart === 'mechanical' ? '기계' : '전기';
  var btn = document.getElementById('btnPart');
  if (btn) btn.innerHTML = '파트: ' + label + ' <span class="arrow">▼</span>';
}

function buildLocDropdown() {
  var locs = getLocations();
  var h = '<div class="dropdown-item' + (curLoc === 'ALL' ? ' active' : '') + '" data-val="ALL">전체</div>';
  for (var i = 0; i < locs.length; i++) {
    var lbl = locs[i].replace(/\n/g, ' ');
    h += '<div class="dropdown-item' + (locs[i] === curLoc ? ' active' : '') + '" data-val="' + esc(locs[i]) + '">' + lbl + '</div>';
  }
  var list = document.getElementById('listLoc');
  list.innerHTML = h;
  updateLocBtnLabel();
  list.onclick = function (e) {
    var item = e.target.closest('.dropdown-item');
    if (!item) return;
    setLoc(item.getAttribute('data-val'));
    document.getElementById('ddLoc').classList.remove('open');
  };
}

function setLoc(lc) {
  curLoc = lc;
  localStorage.setItem('seah_curLoc', lc);
  curMainPage = 1;
  updateLocBtnLabel();
  buildLocDropdown();
  render();
}

function updateLocBtnLabel() {
  var label = curLoc === 'ALL' ? '전체' : curLoc.replace(/\n/g, ' ');
  if (label.length > 10) label = label.substring(0, 10) + '…';
  var btn = document.getElementById('btnLoc');
  btn.innerHTML = '위치: ' + label + ' <span class="arrow">▼</span>';
  if (curLoc !== 'ALL') btn.classList.add('has-value'); else btn.classList.remove('has-value');
}

function buildFreqDropdown() {
  var freqs = [{ v: 'ALL', l: '전체' }, { v: 'D', l: '일 (D)' }, { v: 'W', l: '주 (W)' }, { v: 'M', l: '월 (M)' }];
  var h = '';
  for (var i = 0; i < freqs.length; i++) {
    h += '<div class="dropdown-item' + (freqs[i].v === curFreq ? ' active' : '') + '" data-val="' + freqs[i].v + '">' + freqs[i].l + '</div>';
  }
  var list = document.getElementById('listFreq');
  list.innerHTML = h;
  updateFreqBtnLabel();
  list.onclick = function (e) {
    var item = e.target.closest('.dropdown-item');
    if (!item) return;
    setFreq(item.getAttribute('data-val'));
    document.getElementById('ddFreq').classList.remove('open');
  };
}

function setFreq(f) {
  curFreq = f;
  localStorage.setItem('seah_curFreq', f);
  curMainPage = 1;
  updateFreqBtnLabel();
  buildFreqDropdown();
  render();
}

function updateFreqBtnLabel() {
  var labels = { 'ALL': '전체', 'D': '일 (D)', 'W': '주 (W)', 'M': '월 (M)' };
  var btn = document.getElementById('btnFreq');
  btn.innerHTML = '주기: ' + (labels[curFreq] || '전체') + ' <span class="arrow">▼</span>';
  if (curFreq !== 'ALL') btn.classList.add('has-value'); else btn.classList.remove('has-value');
}

function getAllItems(line, part) {
  var l = line || curLine;
  var p = part || curPart;
  var base = getBaseItems(l, p);
  var ckey = l + '_' + p;
  var custom = customItems[ckey] || [];
  var meta = metaOverrides[ckey] || {};

  var all = base.map(function (it, idx) {
    var m = meta[idx] || {};
    return Object.assign({}, it, m, { baseIndex: idx });
  }).concat(custom.map(function (it, idx) {
    var m = meta[base.length + idx] || {};
    return Object.assign({}, it, m, { customIndex: idx, baseIndex: base.length + idx });
  }));

  return all;
}

function getLocations() {
  var items = getAllItems().filter(function (it) { return !it.hidden; });
  var seen = [];
  for (var i = 0; i < items.length; i++) {
    if (seen.indexOf(items[i].location) === -1) seen.push(items[i].location);
  }
  return seen;
}

function updateEditUI() {
  var eb = document.getElementById('editBtn'), sb = document.getElementById('saveBtn');
  var dh = document.getElementById('delColHead');
  var arb = document.getElementById('addRowBar'), oab = document.getElementById('obsAddBar');
  var br = document.getElementById('batchRow'), bds = document.getElementById('batchDelSpacer');
  var ib = document.getElementById('importBtn');

  if (editing) {
    document.body.classList.add('editing');
    eb.textContent = '수정 취소'; eb.classList.add('editing');
    sb.style.display = '';
    if (dh) dh.style.display = '';
    arb.classList.add('show'); oab.classList.add('show');
    br.style.display = ''; bds.style.display = '';
    if (curLine !== 'OBS' && curLine !== 'DASH') ib.style.display = 'inline-block';
  } else {
    document.body.classList.remove('editing');
    eb.textContent = '수정하기'; eb.classList.remove('editing');
    sb.style.display = 'none';
    if (dh) dh.style.display = 'none';
    arb.classList.remove('show'); oab.classList.remove('show');
    br.style.display = 'none'; bds.style.display = 'none';
    ib.style.display = 'none';
  }
}

function esc(s) { return String(s || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;'); }

function goLine(line, keepFilters) {
  curLine = line;
  localStorage.setItem('seah_curLine', line);
  var btns = document.querySelectorAll('.nav-btn');
  btns.forEach(function (b) {
    b.classList.remove('active');
    if (b.getAttribute('data-line') === line) b.classList.add('active');
  });

  // Close mobile menu
  const sidebar = document.getElementById('sidebar');
  const menuToggle = document.getElementById('menuToggle');
  if (sidebar) sidebar.classList.remove('opened');
  if (menuToggle) menuToggle.classList.remove('active');
  var title = '';
  if (curLine === 'DASH') title = '설비 점검 종합 현황';
  else if (curLine === 'OBS') {
    title = '이상 발견 통합 관리';
  }
  else title = curLine + ' 설비 점검 현황';
  document.getElementById('pageTitle').textContent = title;

  if (!keepFilters) {
    curLoc = 'ALL'; curFreq = 'ALL'; curObsLine = 'ALL'; curObsPage = 1; curMainPage = 1;
    localStorage.setItem('seah_curLoc', 'ALL');
    localStorage.setItem('seah_curFreq', 'ALL');
    localStorage.setItem('seah_curObsLine', 'ALL');
  }

  editing = false; updateEditUI();
  if (curLine !== 'DASH' && curLine !== 'OBS') {
    buildLocDropdown(); buildFreqDropdown();
    fetchWeekData();
  }
  render();
}

function fetchWeekData() {
  if (curLine === 'DASH' || curLine === 'OBS') return;
  
  // 주간 데이터 동기화: 현재 선택된 날짜가 포함된 일주일치 데이터를 한꺼번에 가져옴
  var d = new Date(curDate + 'T00:00:00');
  var day = d.getDay();
  var diff = d.getDate() - day + (day === 0 ? -6 : 1);
  var mon = new Date(d); mon.setDate(diff);
  var sun = new Date(mon); sun.setDate(mon.getDate() + 6);
  
  var formatDate = function(date) {
    return date.getFullYear() + '-' + String(date.getMonth() + 1).padStart(2, '0') + '-' + String(date.getDate()).padStart(2, '0');
  };
  
  var startKey = curLine + '_' + curPart + '_' + formatDate(mon);
  var endKey = curLine + '_' + curPart + '_' + formatDate(sun);
  
  db.collection('inspections')
    .where(firebase.firestore.FieldPath.documentId(), '>=', startKey)
    .where(firebase.firestore.FieldPath.documentId(), '<=', endKey)
    .get()
    .then(querySnapshot => {
      querySnapshot.forEach(doc => {
        saved[doc.id] = doc.data();
      });
      localStorage.setItem('seah_insp', JSON.stringify(saved));
      render();
    }).catch(err => console.error("Weekly Sync Error:", err));
}

function changeObsPage(dir) {
  curObsPage += dir;
  render();
}

function changeMainPage(dir) {
  curMainPage += dir;
  render();
}

function setObsLine(l) {
  curObsLine = l;
  localStorage.setItem('seah_curObsLine', l);
  var btn = document.getElementById('btnObsLine');
  if (btn) btn.innerHTML = '라인: ' + (l === 'ALL' ? '전체' : l) + ' <span class="arrow">▼</span>';
  curObsPage = 1;
  render();
}

function setPeriod(p) {
  curPeriod = p;
  localStorage.setItem('seah_curPeriod', p);
  var btns = document.querySelectorAll('.p-btn');
  btns.forEach(function (b) { b.classList.remove('active'); });
  if (p === 'D' && document.getElementById('pBtnD')) document.getElementById('pBtnD').classList.add('active');
  else if (p === 'W' && document.getElementById('pBtnW')) document.getElementById('pBtnW').classList.add('active');
  else if (p === 'M' && document.getElementById('pBtnM')) document.getElementById('pBtnM').classList.add('active');
  else if (p === 'Y' && document.getElementById('pBtnY')) document.getElementById('pBtnY').classList.add('active');
  else if (p === 'A' && document.getElementById('pBtnA')) document.getElementById('pBtnA').classList.add('active');
  curObsPage = 1;
  render();
}

function isSameWeek(d1, d2) {
  if (!d1 || !d2) return false;
  var date1 = new Date(d1 + 'T00:00:00'), date2 = new Date(d2 + 'T00:00:00');
  var s1 = new Date(date1), s2 = new Date(date2);
  s1.setDate(date1.getDate() - date1.getDay() + (date1.getDay() === 0 ? -6 : 1));
  s2.setDate(date2.getDate() - date2.getDay() + (date2.getDay() === 0 ? -6 : 1));
  return s1.toDateString() === s2.toDateString();
}

function updateDayHeaders() {
  // Find Monday of the week containing curDate (using local time)
  var d = new Date(curDate + 'T00:00:00');
  var day = d.getDay(); // 0:Sun, 1:Mon...
  var diff = d.getDate() - day + (day === 0 ? -6 : 1);
  var mon = new Date(d); mon.setDate(diff);

  var ids = ['thMon', 'thTue', 'thWed', 'thThu', 'thFri', 'thSat', 'thSun'];
  var labels = ['월', '화', '수', '목', '금', '토', '일'];
  for (var i = 0; i < 7; i++) {
    var dd = new Date(mon); dd.setDate(mon.getDate() + i);
    var res = (dd.getMonth() + 1) + '/' + dd.getDate() + ' (' + labels[i] + ')';
    var el = document.getElementById(ids[i]);
    if (el) el.textContent = res;
  }
}

function render() {
  var allItems = getAllItems();
  var key = curLine + '_' + curPart + '_' + curDate;
  updateDayHeaders();

  // 현재 주간의 날짜 라벨 계산 (예: 4/7 (화))
  var weekDateLabels = [];
  var d_head = new Date(curDate + 'T00:00:00');
  var d_shift = d_head.getDay();
  var d_diff = d_head.getDate() - d_shift + (d_shift === 0 ? -6 : 1);
  var mon_head = new Date(d_head); mon_head.setDate(d_diff);
  var dayShortLabels = ['일', '월', '화', '수', '목', '금', '토'];
  for (var k = 0; k < 7; k++) {
    var d_obj = new Date(mon_head); d_obj.setDate(mon_head.getDate() + k);
    weekDateLabels.push((d_obj.getMonth() + 1) + '/' + d_obj.getDate() + ' (' + dayShortLabels[d_obj.getDay()] + ')');
  }

  // Cleanup
  if (document.getElementById('dashView')) document.getElementById('dashView').remove();
  var scrollArea = document.getElementById('scrollArea'), filterArea = document.getElementById('filterArea');
  var progInfo = document.getElementById('progressInfo'), obsSection = document.getElementById('obsSection');
  var editBtn = document.getElementById('editBtn'), exportBtn = document.getElementById('exportBtn');
  var importBtn = document.getElementById('importBtn');
  
  document.getElementById('datePicker').value = curDate;

  // Dashboard Mode
  if (curLine === 'DASH') {
    scrollArea.style.display = 'none'; filterArea.style.display = 'flex'; progInfo.style.display = 'none';
    obsSection.style.display = 'none'; editBtn.style.display = 'none'; exportBtn.style.display = 'none';
    if (importBtn) importBtn.style.display = 'none';
    document.getElementById('ddPart').style.display = 'none';
    document.getElementById('ddLoc').style.display = 'none';
    document.getElementById('ddFreq').style.display = 'none';
    if (document.getElementById('ddObsLine')) document.getElementById('ddObsLine').style.display = 'none';

    var dashPeriodHtml = '<div class="obs-footer-filters" style="margin: 0 1.5rem 1rem 1.5rem"> <span style="font-weight:600; font-size:0.85rem">조회 기간 설정:</span> <div class="period-btns"> <button class="p-btn' + (curPeriod === 'D' ? ' active' : '') + '" onclick="setPeriod(\'D\')">일간</button> <button class="p-btn' + (curPeriod === 'W' ? ' active' : '') + '" onclick="setPeriod(\'W\')">주간</button> <button class="p-btn' + (curPeriod === 'M' ? ' active' : '') + '" onclick="setPeriod(\'M\')">월간</button> <button class="p-btn' + (curPeriod === 'Y' ? ' active' : '') + '" onclick="setPeriod(\'Y\')">연간</button> </div> </div>';

    var lines = ['CPL', 'CRM', 'CGL', '1CCL', '2CCL', '3CCL', 'SSCL'];
    var dh = dashPeriodHtml + '<div class="dash-grid">';

    var curY = curDate.split('-')[0], curMY = curDate.substring(0, 7);
    var keys = Object.keys(saved);

    // Find which day of the week curDate is
    var d_now = new Date(curDate + 'T00:00:00');
    var dayIdx = d_now.getDay(); // 0:Sun, 1:Mon...
    var dayKey = ['sun', 'mon', 'tue', 'wed', 'thu', 'fri', 'sat'][dayIdx];

    var numDaysInPeriod = 1;
    if (curPeriod === 'W') numDaysInPeriod = 7;
    else if (curPeriod === 'M') numDaysInPeriod = new Date(d_now.getFullYear(), d_now.getMonth() + 1, 0).getDate();
    else if (curPeriod === 'Y') numDaysInPeriod = 365;

    for (var i = 0; i < lines.length; i++) {
      var ln = lines[i];
      var mechReq = 0, mechDone = 0;
      var elecReq = 0, elecDone = 0;

      var parts = ['mechanical', 'electrical'];
      for (var pIdx = 0; pIdx < parts.length; pIdx++) {
        var pt = parts[pIdx];
        var items = getAllItems(ln, pt).filter(function (it) { return !it.hidden; });

        for (var j = 0; j < items.length; j++) {
          var item = items[j];
          var freq = item.frequency || 'D';

          var req = 0;
          if (curPeriod === 'D') {
            req = 1;
          } else if (curPeriod === 'W') {
            req = (freq === 'D') ? 7 : 1;
          } else if (curPeriod === 'M') {
            if (freq === 'D') req = numDaysInPeriod;
            else if (freq === 'W') req = Math.ceil(numDaysInPeriod / 7);
            else req = 1;
          } else if (curPeriod === 'Y') {
            if (freq === 'D') req = 365;
            else if (freq === 'W') req = 52;
            else if (freq === 'M') req = 12;
            else req = 1;
          }
          if (req === 0) req = 1;
          if (pt === 'mechanical') mechReq += req; else elecReq += req;

          // Count actual checks across all saved keys of the same week
          var foundDays = {};
          for (var ki = 0; ki < keys.length; ki++) {
            var k = keys[ki];
            var k_parts = k.split('_');
            if (k_parts[0] !== ln || k_parts[1] !== pt) continue;
            var k_date = k_parts[2];
            var match = false;
            if (curPeriod === 'D') match = (k_date === curDate);
            else if (curPeriod === 'W') match = isSameWeek(k_date, curDate);
            else if (curPeriod === 'M') match = k_date.startsWith(curMY);
            else if (curPeriod === 'Y') match = k_date.startsWith(curY);

            if (match) {
              var row = (saved[k].rows || [])[j] || {};
              if (curPeriod === 'D') {
                if (row[dayKey]) foundDays[dayKey] = 1;
              } else {
                if (row.mon) foundDays['mon'] = 1; if (row.tue) foundDays['tue'] = 1;
                if (row.wed) foundDays['wed'] = 1; if (row.thu) foundDays['thu'] = 1;
                if (row.fri) foundDays['fri'] = 1; if (row.sat) foundDays['sat'] = 1;
                if (row.sun) foundDays['sun'] = 1;
              }
            }
          }
          var actual = Object.keys(foundDays).length;
          if (pt === 'mechanical') mechDone += Math.min(actual, req); else elecDone += Math.min(actual, req);
        }
      }

      var m_perc = mechReq > 0 ? Math.round((mechDone / mechReq) * 100) : 0;
      var e_perc = elecReq > 0 ? Math.round((elecDone / elecReq) * 100) : 0;

      dh += '<div class="line-card" onclick="goLine(\'' + ln + '\')">';
      dh += '<div class="card-title"><span>' + ln + '</span></div>';

      dh += '<div class="card-stat" style="display:flex; justify-content:space-between; margin-bottom:4px; font-size:0.75rem"><span>기계</span> <span style="font-weight:700; color:var(--primary)">' + m_perc + '%</span></div>';
      dh += '<div class="prog-bar-bg" style="height:6px; margin-bottom:12px"><div class="prog-bar-fill" style="width:' + m_perc + '%"></div></div>';

      dh += '<div class="card-stat" style="display:flex; justify-content:space-between; margin-bottom:4px; font-size:0.75rem"><span>전기</span> <span style="font-weight:700; color:#10b981">' + e_perc + '%</span></div>';
      dh += '<div class="prog-bar-bg" style="height:6px; margin-bottom:6px"><div class="prog-bar-fill" style="width:' + e_perc + '%; background:#10b981"></div></div>';

      dh += '<div class="card-stat" style="font-size:0.7rem; color:#94a3b8; margin-top:8px">기계: ' + mechDone + '/' + mechReq + ' | 전기: ' + elecDone + '/' + elecReq + '</div>';
      dh += '</div>';
    }
    dh += '</div>';

    // Recent Observations Section
    var allObs = [];
    var allObsKeys = Object.keys(saved).filter(function (k) { return k.split('_').length >= 3; })
      .sort(function (a, b) { return b.split('_')[2].localeCompare(a.split('_')[2]); });
    for (var k_idx = 0; k_idx < allObsKeys.length; k_idx++) {
      var k_obs = allObsKeys[k_idx];
      var parts_obs = k_obs.split('_');
      var date_obs = parts_obs[2];
      var line_obs = parts_obs[0];
      if (line_obs === 'OBS') continue;
      var arr = saved[k_obs].observations || [];
      for (var x_obs = 0; x_obs < arr.length; x_obs++) {
        if (!arr[x_obs].equipment && !arr[x_obs].detail) continue;
        allObs.push({ date: date_obs, line: line_obs, equip: arr[x_obs].equipment, content: arr[x_obs].detail });
      }
      if (allObs.length >= 10) break;
    }

    dh += '<div class="dash-section-title">최근 이상 발견 현황</div>';
    dh += '<div class="dash-obs-area">';
    dh += '<table style="width:100%; border-collapse: collapse;">';
    dh += '<thead><tr style="background:#f8fafc; border-bottom:2px solid var(--border);">';
    dh += '<th style="width:140px; padding:1rem; text-align:center;">날짜</th>';
    dh += '<th style="width:100px; padding:1rem; text-align:center;">라인</th>';
    dh += '<th style="width:300px; padding:1rem; text-align:center;">설비명</th>';
    dh += '<th style="padding:1rem 2rem; text-align:center;">이상 내역</th>';
    dh += '</tr></thead>';
    dh += '<tbody>';
    if (allObs.length === 0) {
      dh += '<tr><td colspan="4" style="color:var(--text-muted); padding:4rem; text-align:center;">최근 기록된 이상 항목이 없습니다.</td></tr>';
    } else {
      for (var o_idx = 0; o_idx < allObs.length && o_idx < 5; o_idx++) {
        var o = allObs[o_idx];
        dh += '<tr style="border-bottom:1px solid var(--border);">';
        dh += '<td style="padding:1.2rem 1rem; text-align:center; color:var(--text-muted);">' + o.date + '</td>';
        dh += '<td style="padding:1.2rem 1rem; text-align:center;"><span class="freq-badge freq-D" style="background:#f1f5f9; color:#475569; font-size:0.75rem; padding:0.3rem 0.6rem">' + o.line + '</span></td>';
        dh += '<td style="padding:1.2rem 1rem; text-align:center; font-weight:600; font-size:0.95rem">' + esc(o.equip) + '</td>';
        dh += '<td style="padding:1.2rem 2rem; white-space:normal; line-height:1.6; font-size:0.9rem; text-align:center;">' + esc(o.content) + '</td>';
        dh += '</tr>';
      }
    }
    dh += '</tbody></table></div>';

    scrollArea.insertAdjacentHTML('afterend', '<div id="dashView">' + dh + '</div>');
    return;
  }


  // Line or OBS Mode Prep
  editBtn.style.display = 'inline-block'; exportBtn.style.display = 'inline-block'; filterArea.style.display = 'flex';
  importBtn.style.display = (editing && curLine !== 'OBS' && curLine !== 'DASH') ? 'inline-block' : 'none';
  var allItems = getAllItems();
  var key = curLine + '_' + curPart + '_' + curDate;
  var sd = saved[key] || { rows: [], observations: [] };

  // OBS Mode
  if (curLine === 'OBS') {
    scrollArea.style.display = 'none'; progInfo.style.display = 'none'; obsSection.style.display = 'block';
    document.getElementById('ddPart').style.display = 'none'; document.getElementById('ddLoc').style.display = 'none'; document.getElementById('ddFreq').style.display = 'none';
    document.getElementById('obsHeadNorm').style.display = 'none';
    document.getElementById('obsHeadGlobal').style.display = '';

    if (!document.getElementById('ddObsLine')) {
      var div = document.createElement('div');
      div.className = 'dropdown'; div.id = 'ddObsLine';
      div.innerHTML = '<button class="dropdown-btn" id="btnObsLine">라인: 전체 <span class="arrow">▼</span></button><div class="dropdown-list" id="listObsLine"></div>';
      filterArea.insertBefore(div, filterArea.firstChild);
      buildObsFilter();
    } else document.getElementById('ddObsLine').style.display = 'inline-block';

    curSearch = '';
    var filterLines = (curObsLine === 'ALL') ? ['CPL', 'CRM', 'CGL', '1CCL', '2CCL', '3CCL', 'SSCL'] : [curObsLine];
    var dataRowsFinal = [];
    var keys = Object.keys(saved).sort(function (a, b) { return b.split('_')[2].localeCompare(a.split('_')[2]); });
    var curY = curDate.split('-')[0], curMY = curDate.substring(0, 7);

    for (var ki = 0; ki < keys.length; ki++) {
      var k = keys[ki];
      var parts = k.split('_');
      if (parts.length < 3) continue;
      var line = parts[0], part = parts[1], date = parts[2];
      if (filterLines.indexOf(line) === -1) continue;
      var match = (curPeriod === 'D' && date === curDate) ||
        (curPeriod === 'W' && isSameWeek(date, curDate)) ||
        (curPeriod === 'M' && date.startsWith(curMY)) ||
        (curPeriod === 'Y' && date.startsWith(curY)) ||
        (curPeriod === 'A');
      if (match) {
        var oarr = (saved[k] || {}).observations || [];
        for (var i = 0; i < oarr.length; i++) {
          var obj = oarr[i];
          if (!obj.equipment && !obj.detail && !editing) continue;

          // [Robust Validation]
          if ((obj.detail || "").startsWith("[자동]")) {
            var norm = function (s) { return String(s || "").replace(/\s+/g, ""); };
            var master = getBaseItems(line, part).concat(customItems[line + '_' + part] || []);
            var isValid = master.some(function (it) {
              return norm(it.location) === norm(obj.location) && norm(it.equipment) === norm(obj.equipment);
            });
            if (!isValid) continue;
          }

          // Global Deduplication in UI
          var isDup = dataRowsFinal.some(function (existing) {
            return existing.date === date &&
              existing.line === line &&
              (existing.obj.location || "") === (obj.location || "") &&
              existing.obj.equipment === obj.equipment &&
              existing.obj.detail === obj.detail;
          });

          if (!isDup) {
            dataRowsFinal.push({ date: date, line: line, part: part, idx: i, obj: obj });
          }
        }
      }
    }

    // Pagination Logic
    var totalItems = dataRowsFinal.length;
    var totalPages = Math.ceil(totalItems / 10) || 1;
    if (curObsPage > totalPages) curObsPage = totalPages;
    var dataRows = dataRowsFinal.slice((curObsPage - 1) * 10, curObsPage * 10);

    var oh = '';
    var dayNames = ['sun', 'mon', 'tue', 'wed', 'thu', 'fri', 'sat'];

    for (var i = 0; i < dataRows.length; i++) {
      var d = dataRows[i];

      // 1. Find Inspection Photos
      var k_insp = d.line + '_' + d.part + '_' + d.date;
      var dInp = new Date(d.date + 'T00:00:00');
      var dayKey = dayNames[dInp.getDay()];
      var inspPhotos = [];

      // To find inspection photos, we need the index of the equipment in the BASE + CUSTOM list
      var allLineItems = getBaseItems(d.line, d.part).concat(customItems[d.line + '_' + d.part] || []);
      var origIdx = -1; 
      // Robust matching to find origIdx even with whitespace differences
      var norm = function(s) { return String(s || "").replace(/\s+/g, ""); };
      for (var j = 0; j < allLineItems.length; j++) {
        var it = allLineItems[j];
        if (norm(it.location) === norm(d.obj.location) && norm(it.equipment) === norm(d.obj.equipment)) {
          origIdx = j; break;
        }
      }

          // 1. Find Inspection Photos - Search across the whole week since data is spread
          var wDates = getWeekDates(d.date);
          for (var wd = 0; wd < 7; wd++) {
            var wk = d.line + '_' + d.part + '_' + wDates[wd];
            if (origIdx !== -1 && saved[wk] && saved[wk].rows && saved[wk].rows[origIdx]) {
              var r = saved[wk].rows[origIdx];
              if (r.photos && r.photos[dayKey] && r.photos[dayKey].length > 0) {
                inspPhotos = r.photos[dayKey];
                k_insp = wk; // Important: use the key that actually has the photos
                break;
              }
            }
          }
          
          // 2. Action Photos
      var actionPhotos = d.obj.actionPhotos || [];

      oh += '<tr data-line="' + d.line + '" data-part="' + d.part + '" data-date="' + d.date + '" data-oidx="' + d.idx + '" data-action-photos="' + (d.obj.actionPhotos || []).join('|') + '">';
      oh += '<td style="font-size:0.75rem; color:#64748b">' + d.date + '</td><td style="background:#f1f5f9; font-weight:600">' + d.line + '</td>';
      oh += '<td><input class="obs-inp" style="width:70px" ' + (editing ? '' : 'disabled') + ' value="' + esc(d.obj.location || "") + '" placeholder="위치"></td>';
      oh += '<td><input class="obs-inp" ' + (editing ? '' : 'disabled') + ' value="' + esc(d.obj.equipment) + '" placeholder="설비명"></td>';
      oh += '<td><input class="obs-inp" ' + (editing ? '' : 'disabled') + ' value="' + esc(d.obj.detail) + '" placeholder="내용"></td>';
      oh += '<td><input class="obs-inp" ' + (editing ? '' : 'disabled') + ' value="' + esc(d.obj.confirm) + '" placeholder="조치 여부"></td>';

      // Anomaly Photo Button
      var pBtnClass = inspPhotos.length > 0 ? 'photo-badge-btn has-photo' : 'photo-badge-btn';
      oh += '<td><button class="' + pBtnClass + '" onclick="openPhotoModal(\'' + k_insp + '\',' + origIdx + ',\'' + dayKey + '\',\'' + d.line + '\',\'' + d.part + '\')">📸 <span class="badge' + (inspPhotos.length === 0 ? ' empty' : '') + '">' + inspPhotos.length + '</span></button></td>';

      // Action Photo Button
      var aBtnClass = actionPhotos.length > 0 ? 'photo-badge-btn has-photo' : 'photo-badge-btn';
      oh += '<td><button class="' + aBtnClass + '" onclick="openActionPhotoModal(\'' + d.line + '\',\'' + d.part + '\',\'' + d.date + '\',' + d.idx + ')">🛠️ <span class="badge' + (actionPhotos.length === 0 ? ' empty' : '') + '">' + actionPhotos.length + '</span></button></td>';

      oh += '<td class="obs-del-col">' + (editing ? '<button class="btn-del" onclick="deleteObsRowGlobal(\'' + d.line + '\',\'' + d.part + '\',\'' + d.date + '\',' + d.idx + ')">삭제</button>' : '') + '</td></tr>';
    }

    // REMOVED: Auto-generating 5 empty rows logic to prevent confusion with non-persistent data
    var ptxt = (curPeriod === 'D' ? '오늘(일)' : curPeriod === 'M' ? '이번 달(월)' : curPeriod === 'Y' ? '올해(년)' : '전체 이력');
    document.getElementById('obsGlobalTitle').textContent = '이상 발견 통합 관리';
    if (!oh && !editing) oh = '<tr><td colspan="6" style="padding:4rem; color:#94a3b8">내역이 없습니다. (조회 기간: ' + ptxt + ')</td></tr>';
    document.getElementById('obody').innerHTML = oh;

    // Add Pagination UI below table
    var pagHtml = '';
    if (totalPages > 1) {
      pagHtml += '<div class="pagination" style="display:flex; align-items:center; justify-content:center; gap:1rem; margin-top:1.5rem; padding:1rem; border-top:1px solid var(--border);">';
      pagHtml += '<button class="btn" ' + (curObsPage <= 1 ? 'disabled style="opacity:0.5; cursor:default"' : 'onclick="changeObsPage(-1)"') + '>이전</button>';
      pagHtml += '<span style="font-weight:700; color:var(--text-main); font-size:0.9rem">' + curObsPage + ' / ' + totalPages + '</span>';
      pagHtml += '<button class="btn" ' + (curObsPage >= totalPages ? 'disabled style="opacity:0.5; cursor:default"' : 'onclick="changeObsPage(1)"') + '>다음</button>';
      pagHtml += '</div>';
    }
    // Remove existing pagination if any
    var existingPag = document.getElementById('obsPagination');
    if (existingPag) existingPag.remove();
    var pDiv = document.createElement('div');
    pDiv.id = 'obsPagination';
    pDiv.innerHTML = pagHtml;
    obsSection.appendChild(pDiv);

    return;
  }

  // Per-line View
  scrollArea.style.display = 'block'; progInfo.style.display = 'flex'; obsSection.style.display = 'none';
  if (document.getElementById('ddObsLine')) document.getElementById('ddObsLine').style.display = 'none';
  document.getElementById('ddPart').style.display = 'inline-block';
  document.getElementById('ddLoc').style.display = 'inline-block';
  document.getElementById('ddFreq').style.display = 'inline-block';

  var items = allItems.filter(function (it) { return !it.hidden; });
  if (curLoc !== 'ALL') items = items.filter(function (it) { return it.location === curLoc; });
  if (curFreq !== 'ALL') items = items.filter(function (it) { return (it.frequency || 'D') === curFreq; });

  // Merge data from all documents of the same week for better persistence
  // Sort keys to ensure newer documents (by date) overwrite older ones during merge
  var mergedRows = [];
  var keys = Object.keys(saved).sort(function (a, b) {
    var d1 = a.split('_')[2] || '0', d2 = b.split('_')[2] || '0';
    return d1.localeCompare(d2);
  });

  for (var ki = 0; ki < keys.length; ki++) {
    var k = keys[ki];
    var k_parts = k.split('_');
    if (k_parts[0] === curLine && k_parts[1] === curPart && isSameWeek(k_parts[2], curDate)) {
      var rows = saved[k].rows || [];
      for (var r_idx = 0; r_idx < rows.length; r_idx++) {
        if (!mergedRows[r_idx]) mergedRows[r_idx] = {};
        var r = rows[r_idx];
        if (!r) continue;

        // Only overwrite if the new data is not empty, to preserve week-level persistence
        var fields = ['mon', 'tue', 'wed', 'thu', 'fri', 'sat', 'sun', 'criteria', 'remarks', 'weekLabel'];
        fields.forEach(function (f) {
          if (r[f] !== undefined && r[f] !== "" && r[f] !== null) mergedRows[r_idx][f] = r[f];
          else if (mergedRows[r_idx][f] === undefined) mergedRows[r_idx][f] = r[f] || "";
        });

        if (r.photos) {
          if (!mergedRows[r_idx].photos) mergedRows[r_idx].photos = {};
          for (var d in r.photos) {
            if (!mergedRows[r_idx].photos[d]) mergedRows[r_idx].photos[d] = [];
            var combined = (mergedRows[r_idx].photos[d] || []).concat(r.photos[d] || []);
            mergedRows[r_idx].photos[d] = Array.from(new Set(combined)).filter(function(p){ return p && p.trim() !== ""; });
          }
        }
      }
    }
  }

  var d_today = new Date(curDate + 'T00:00:00');
  var tDayKey = ['sun', 'mon', 'tue', 'wed', 'thu', 'fri', 'sat'][d_today.getDay()];
  var curWeekNum = getWeekOfMonth(curDate);
  var done = 0;

  for (var i = 0; i < items.length; i++) {
    var idx = allItems.indexOf(items[i]);
    var r = mergedRows[idx] || {};
    // Regardless of cycle, for "Today's Status" header, we check if today's column is filled
    if (r[tDayKey]) done++;
  }

  document.getElementById('statTotal').textContent = items.length;
  document.getElementById('statDone').textContent = done;
  document.getElementById('statPending').textContent = items.length - done;
  document.getElementById('progBarFill').style.width = (items.length > 0 ? (done / items.length * 100) : 0) + '%';

  // Pagination for main table
  var totalItems = items.length;
  var totalPages = Math.ceil(totalItems / 10) || 1;
  if (curMainPage > totalPages) curMainPage = totalPages;
  var pagedItems = items.slice((curMainPage - 1) * 10, curMainPage * 10);

  var h = '', freqLabel = { 'D': '일', 'W': '주', 'M': '월' };
  for (var i = 0; i < pagedItems.length; i++) {
    var it = pagedItems[i], origIdx = it.baseIndex, r = mergedRows[origIdx] || {}, freq = it.frequency || 'D';
    var seqNum = (curMainPage - 1) * 10 + (i + 1);
    h += '<tr data-idx="' + origIdx + '">';
    if (editing) {
      h += '<td>' + seqNum + '</td><td><input class="edit-inp" data-field="location" value="' + esc(it.location) + '"></td><td><input class="edit-inp" data-field="equipment" value="' + esc(it.equipment) + '"></td><td><select class="edit-sel" data-field="frequency"><option value="D"' + (freq === 'D' ? ' selected' : '') + '>일</option><option value="W"' + (freq === 'W' ? ' selected' : '') + '>주</option><option value="M"' + (freq === 'M' ? ' selected' : '') + '>월</option></select></td>';
    } else {
      h += '<td>' + seqNum + '</td><td class="loc-cell">' + it.location + '</td><td class="eq-cell" style="text-align:left">' + it.equipment + '</td><td><span class="freq-badge freq-' + freq + '">' + (freqLabel[freq] || freq) + '</span></td>';
    }

    // 시기(Timing) Column
    var weekVal = r.weekLabel || it.weekLabel || '';
    h += '<td><input class="inp week-inp" ' + (editing ? '' : 'disabled') + ' value="' + esc(weekVal) + '" placeholder="시기" data-field="weekLabel"></td>';

    // Criteria Column (Master criteria should take precedence)
    var critVal = it.criteria || '';
    h += '<td><input class="inp crit-inp" ' + (editing ? '' : 'disabled') + ' value="' + esc(critVal) + '" placeholder="기준" data-field="criteria"></td>';

    var days = ['mon', 'tue', 'wed', 'thu', 'fri', 'sat', 'sun'];

    // 주/월 단위 점검 항목 또는 '시기'가 입력된 항목의 경우, 이번 주 내에 완료된 날짜가 있는지 확인
    var compInfo = null;
    if (freq === 'W' || freq === 'M' || weekVal) {
      for (var d = 0; d < 7; d++) {
        if (r[days[d]]) {
          compInfo = { dateLabel: weekDateLabels[d], val: r[days[d]] };
          break;
        }
      }
    }

    for (var d = 0; d < days.length; d++) {
      var dayKey = days[d];
      var val = r[dayKey] || '';

      var displayVal = val;
      var statusClass = '';
      var isDisabled = !editing;

      if (!editing && !val && compInfo) {
        displayVal = compInfo.dateLabel + ' 완료';
        statusClass = ' status-autofill';
        isDisabled = true;
      } else {
        var status = checkValueStatus(critVal, val);
        if (status === 'ok') statusClass = ' status-ok';
        else if (status === 'error') statusClass = ' status-error';
        
        // 주/월 점검 항목의 경우, 수정 모드에서도 이미 다른 칸에 입력이 있으면 이 빈 칸은 비활성화
        if (editing && (freq === 'W' || freq === 'M') && !val && compInfo) {
           isDisabled = true;
           statusClass += ' status-locked';
        }
      }

      var pArr = (r.photos && r.photos[dayKey]) ? r.photos[dayKey].filter(function(p){ return p && p.trim() !== ""; }) : [];
      var photoCount = pArr.length;

      h += '<td class="day-col"><div class="day-cell-wrapper"><input class="inp' + statusClass + '" ' + (isDisabled ? 'disabled' : '') + ' value="' + esc(displayVal) + '" placeholder="-">';

      if (photoCount > 0 || editing) {
        var icon = photoCount > 0 ? '<svg viewBox="0 0 24 24"><path d="M4 4h3l2-2h6l2 2h3a2 2 0 0 1 2 2v12a2 2 0 0 1-2 2H4a2 2 0 0 1-2-2V6a2 2 0 0 1 2-2z"></path><circle cx="12" cy="13" r="4"></circle></svg>' : '➕';
        var btnClass = photoCount > 0 ? '' : ' empty-photo';
        // 이 칸이 비활성화된 경우(다른 칸이 이미 입력됨) 사진 버튼도 숨김(신규 추가 방지)
        if (!isDisabled || val) {
          h += '<div class="photo-btn-wrap"><button class="photo-btn' + btnClass + '" onclick="openPhotoModal(\'' + key + '\',' + origIdx + ',\'' + dayKey + '\')">' + icon + '</button></div>';
        }
      }
      h += '</div></td>';
    }
    h += '<td><input class="rem" ' + (editing ? '' : 'disabled') + ' value="' + esc(r.remarks || '') + '" title="' + esc(r.remarks || '') + '" placeholder="비고"></td><td style="display:' + (editing ? '' : 'none') + '"><button class="btn-del" data-idx="' + origIdx + '">삭제</button></td></tr>';
  }
  document.getElementById('tbody').innerHTML = h;
  document.getElementById('tbody').querySelectorAll('tr').forEach(function (tr) {
    var cInp = tr.querySelector('input[data-field="criteria"]');
    var dInps = tr.querySelectorAll('.day-col .inp');

    var updateRowColors = function () {
      var cv = cInp.value.trim();
      var freqSel = tr.querySelector('.edit-sel');
      var freq = freqSel ? freqSel.value : 'D';
      
      var hasAnyValue = Array.from(dInps).some(function(di) { return di.value.trim() !== ""; });

      dInps.forEach(function (di) {
        var dv = di.value.trim();
        di.classList.remove('status-ok', 'status-error');
        
        var status = checkValueStatus(cv, dv);
        if (status === 'ok') di.classList.add('status-ok');
        else if (status === 'error') di.classList.add('status-error');

        // 주/월 점검 실시간 잠금 로직
        if (freq === 'W' || freq === 'M') {
          if (hasAnyValue && dv === "") {
            di.disabled = true;
            di.classList.add('status-locked');
          } else {
            di.disabled = false;
            di.classList.remove('status-locked');
          }
        }
      });
    };

    cInp.oninput = updateRowColors;
    dInps.forEach(function (di) { di.oninput = updateRowColors; });
    var fSel = tr.querySelector('.edit-sel');
    if (fSel) fSel.onchange = updateRowColors;
  });

  // Add Pagination UI for main table
  var pagHtml = '';
  if (totalPages > 1) {
    pagHtml += '<div class="pagination" style="display:flex; align-items:center; justify-content:center; gap:1rem; padding:1rem; border-top:1px solid var(--border); background:#fcfcfc">';
    pagHtml += '<button class="btn" ' + (curMainPage <= 1 ? 'disabled style="opacity:0.5; cursor:default"' : 'onclick="changeMainPage(-1)"') + '>이전</button>';
    pagHtml += '<span style="font-weight:700; color:var(--text-main); font-size:0.9rem">' + curMainPage + ' / ' + totalPages + '</span>';
    pagHtml += '<button class="btn" ' + (curMainPage >= totalPages ? 'disabled style="opacity:0.5; cursor:default"' : 'onclick="changeMainPage(1)"') + '>다음</button>';
    pagHtml += '</div>';
  }
  var existingPag = document.getElementById('mainPagination');
  if (existingPag) existingPag.remove();
  var pDiv = document.createElement('div');
  pDiv.id = 'mainPagination';
  pDiv.innerHTML = pagHtml;
  scrollArea.appendChild(pDiv);

  scrollArea.scrollTop = 0;
}

function buildObsLineDropdown() {
  var lines = ['ALL', 'CPL', 'CRM', 'CGL', '1CCL', '2CCL', '3CCL', 'SSCL'];
  var list = document.getElementById('listObsLine');
  if (!list) return;
  var h = '';
  for (var i = 0; i < lines.length; i++) {
    h += '<div class="dropdown-item' + (curObsLine === lines[i] ? ' active' : '') + '" onclick="setObsLine(\'' + lines[i] + '\')">' + (lines[i] === 'ALL' ? '전체 라인' : lines[i]) + '</div>';
  }
  list.innerHTML = h;
  document.getElementById('btnObsLine').innerHTML = '라인: ' + (curObsLine === 'ALL' ? '전체' : curObsLine) + ' <span class="arrow">▼</span>';
}
function setObsLine(ln) {
  curObsLine = ln;
  localStorage.setItem('seah_curObsLine', ln);
  buildObsLineDropdown();
  render();
}

function closePwdModal() {
  document.getElementById('pwdModal').classList.remove('show');
}

async function confirmPwd() {
  const pwdInp = document.getElementById('editPwdInp');
  const pwd = pwdInp.value;
  if (!pwd) return;

  try {
    // Firestore에서 설정을 가져옴 (collection: settings, doc: security)
    const doc = await db.collection('settings').doc('security').get();
    let correctPwd = '2026'; // 기본값 (문서가 없을 경우)
    
    if (doc.exists && doc.data().editPassword) {
      correctPwd = doc.data().editPassword;
    }

    if (pwd === correctPwd) {
      sessionStorage.setItem('seah_authenticated', 'true');
      savedBackup = JSON.stringify(saved);
      customBackup = JSON.stringify(customItems);
      metaBackup = JSON.stringify(metaOverrides);
      editing = true;
      closePwdModal();
      updateEditUI();
      render();
    } else {
      alert('비밀번호가 틀렸습니다.');
      pwdInp.value = '';
      pwdInp.focus();
    }
  } catch (err) {
    console.error("Auth Error:", err);
    // 네트워크 오류 시 로컬 하드코딩 비상용 (선택 사항)
    if (pwd === '2026') { 
       /* 비상 로그인 허용로직 생략 가능 */ 
    }
    alert('인증 서버에 접속할 수 없습니다. 네트워크를 확인해주세요.');
  }
}

function addItem() {
  var loc = document.getElementById('newLoc').value.trim();
  var equip = document.getElementById('newEquip').value.trim();
  var freq = document.getElementById('newFreq').value;
  if (!loc || !equip) { alert('위치와 설비명을 입력해주세요.'); return; }
  var ckey = curLine + '_' + curPart;
  if (!customItems[ckey]) customItems[ckey] = [];
  customItems[ckey].push({ location: loc, equipment: equip, frequency: freq, criteria: '' });
  localStorage.setItem('seah_custom', JSON.stringify(customItems));
  document.getElementById('addModal').classList.remove('show');
  buildLocDropdown(); render();
}

function deleteItem(origIdx) {
  if (!confirm('이 항목을 삭제하시겠습니까?')) return;
  var base = getBaseItems(curLine, curPart);
  var ckey = curLine + '_' + curPart;
  var custom = customItems[ckey] || [];

  if (origIdx < base.length) {
    if (!metaOverrides[ckey]) metaOverrides[ckey] = {};
    metaOverrides[ckey][origIdx] = Object.assign({}, metaOverrides[ckey][origIdx] || base[origIdx], { hidden: true });
    localStorage.setItem('seah_meta', JSON.stringify(metaOverrides));
    db.collection('settings').doc('metaOverrides').set(metaOverrides).catch(e => { });
    render();
    return;
  }
  var customIdx = origIdx - base.length;
  custom.splice(customIdx, 1);
  customItems[ckey] = custom;

  // 인덱스가 변경되었으므로 메모리에 저장된 모든 시점의 점검 데이터도 인덱스를 밀어줍니다.
  for (var k in saved) {
    if (saved[k].rows) {
      saved[k].rows.splice(origIdx, 1);
    }
  }

  localStorage.setItem('seah_insp', JSON.stringify(saved));
  localStorage.setItem('seah_custom', JSON.stringify(customItems));
  db.collection('settings').doc('customItems').set(customItems).catch(e => { });
  buildLocDropdown(); render();
}

function addObsRow() {
  if (curLine === 'OBS') {
    var line = curObsLine === 'ALL' ? 'CPL' : curObsLine;
    var key = line + '_mechanical_' + curDate;
    var sd = saved[key] || { rows: [], observations: [] };
    if (!sd.observations) sd.observations = [];
    sd.observations.push({ location: '', equipment: '', detail: '', confirm: '' });
    saved[key] = sd;
  } else {
    var key = curLine + '_' + curPart + '_' + curDate;
    var sd = saved[key] || { rows: [], observations: [] };
    if (!sd.observations) sd.observations = [];
    sd.observations.push({ location: '', equipment: '', detail: '', confirm: '' });
    saved[key] = sd;
  }
  render();
}

function deleteObsRowGlobal(line, part, date, idx) {
  if (!confirm('이 항목을 삭제하시겠습니까?')) return;
  var k = line + '_' + part + '_' + date;
  if (saved[k] && saved[k].observations) {
    saved[k].observations.splice(idx, 1);
    localStorage.setItem('seah_insp', JSON.stringify(saved));
    db.collection('inspections').doc(k).set(saved[k]).then(() => {
      render();
    });
  }
}

function buildObsFilter() {
  var lines = ['ALL', 'CPL', 'CRM', 'CGL', '1CCL', '2CCL', '3CCL', 'SSCL'];
  var h = '';
  for (var i = 0; i < lines.length; i++) {
    h += '<div class="dropdown-item' + (lines[i] === curObsLine ? ' active' : '') + '" data-val="' + lines[i] + '">' + (lines[i] === 'ALL' ? '전체' : lines[i]) + '</div>';
  }
  var list = document.getElementById('listObsLine');
  list.innerHTML = h;
  document.getElementById('btnObsLine').innerHTML = '라인: ' + (curObsLine === 'ALL' ? '전체' : curObsLine) + ' <span class="arrow">▼</span>';

  list.onclick = function (e) {
    var item = e.target.closest('.dropdown-item');
    if (!item) return;
    curObsLine = item.getAttribute('data-val');
    document.getElementById('ddObsLine').classList.remove('open');
    buildObsFilter();
    render();
  }

  document.getElementById('btnObsLine').onclick = function (e) {
    e.stopPropagation();
    document.getElementById('ddObsLine').classList.toggle('open');
  }
}

function deleteObsRow(idx) {
  if (!confirm('이 이상 항목을 삭제하시겠습니까?')) return;
  var key = curLine + '_' + curPart + '_' + curDate;
  var sd = saved[key] || { rows: [], observations: [] };
  if (sd.observations && idx < sd.observations.length) {
    sd.observations.splice(idx, 1);
    saved[key] = sd;
    localStorage.setItem('seah_insp', JSON.stringify(saved));
    render();
  }
}

function bulkNormal(day) {
  var isCrit = (day === 'criteria');
  var valToSet = '양호';

  if (isCrit) {
    valToSet = prompt('일괄 입력할 점검 기준을 입력해주세요. (예: 70~85bar)', '양호');
    if (valToSet === null) return;
  } else {
    if (!confirm('현재 요일의 모든 항목을 \'양호\'로 일괄 입력하시겠습니까?')) return;
  }

  var allItems = getAllItems();
  var items = allItems;
  if (curLoc !== 'ALL') items = items.filter(function (it) { return it.location === curLoc; });
  if (curFreq !== 'ALL') items = items.filter(function (it) { return (it.frequency || 'D') === curFreq; });

  var key = curLine + '_' + curPart + '_' + curDate;
  var ckey = curLine + '_' + curPart;
  var base = getBaseItems(curLine, curPart);

  if (!saved[key]) saved[key] = { rows: [] };
  if (!saved[key].rows) saved[key].rows = [];

  for (var i = 0; i < items.length; i++) {
    var origIdx = items[i].baseIndex;
    if (!saved[key].rows[origIdx]) saved[key].rows[origIdx] = {};
    saved[key].rows[origIdx][day] = valToSet;

    // 점검 기준(Master)도 전역적으로 업데이트
    if (isCrit) {
      if (origIdx >= base.length) {
        var cIdx = origIdx - base.length;
        if (customItems[ckey] && customItems[ckey][cIdx]) customItems[ckey][cIdx].criteria = valToSet;
      } else {
        if (!metaOverrides[ckey]) metaOverrides[ckey] = {};
        var existing = metaOverrides[ckey][origIdx] || base[origIdx];
        metaOverrides[ckey][origIdx] = Object.assign({}, existing, { criteria: valToSet });
      }
    }
  }

  localStorage.setItem('seah_insp', JSON.stringify(saved));
  db.collection('inspections').doc(key).set(saved[key]);

  if (isCrit) {
    localStorage.setItem('seah_custom', JSON.stringify(customItems));
    localStorage.setItem('seah_meta', JSON.stringify(metaOverrides));
    db.collection('settings').doc('customItems').set(customItems).catch(e => { });
    db.collection('settings').doc('metaOverrides').set(metaOverrides).catch(e => { });
  }

  render();
}

function bulkClear(day) {
  var msg = day === 'ALL' ? '현재 화면의 모든 내용을 모두 비우시겠습니까?' : (day === 'criteria' ? '모든 항목의 기준을 모두 비우시겠습니까?' : '해당 요일의 내용을 모두 비우시겠습니까?');
  if (!confirm(msg)) return;

  var allItems = getAllItems();
  var items = allItems;
  if (curLoc !== 'ALL') items = items.filter(function (it) { return it.location === curLoc; });
  if (curFreq !== 'ALL') items = items.filter(function (it) { return (it.frequency || 'D') === curFreq; });

  var keys = Object.keys(saved);
  var days = day === 'ALL' ? ['mon', 'tue', 'wed', 'thu', 'fri', 'sat', 'sun', 'criteria'] : [day];
  var isCrit = days.includes('criteria');
  var ckey = curLine + '_' + curPart;
  var base = getBaseItems(curLine, curPart);

  // Clear from all documents in the same week
  for (var ki = 0; ki < keys.length; ki++) {
    var k = keys[ki];
    var kp = k.split('_');
    if (kp[0] === curLine && kp[1] === curPart && isSameWeek(kp[2], curDate)) {
      if (!saved[k].rows) continue;
      for (var i = 0; i < items.length; i++) {
        var origIdx = items[i].baseIndex;
        if (saved[k].rows[origIdx]) {
          for (var d = 0; d < days.length; d++) {
            saved[k].rows[origIdx][days[d]] = '';
          }
        }
      }
      db.collection('inspections').doc(k).set(saved[k]);
    }
  }

  // 기준(Master)도 전역적으로 비우기
  if (isCrit) {
    for (var i = 0; i < items.length; i++) {
      var origIdx = items[i].baseIndex;
      if (origIdx >= base.length) {
        var cIdx = origIdx - base.length;
        if (customItems[ckey] && customItems[ckey][cIdx]) customItems[ckey][cIdx].criteria = '';
      } else {
        if (!metaOverrides[ckey]) metaOverrides[ckey] = {};
        var existing = metaOverrides[ckey][origIdx] || base[origIdx];
        metaOverrides[ckey][origIdx] = Object.assign({}, existing, { criteria: '' });
      }
    }
    localStorage.setItem('seah_custom', JSON.stringify(customItems));
    localStorage.setItem('seah_meta', JSON.stringify(metaOverrides));
    db.collection('settings').doc('customItems').set(customItems).catch(e => { });
    db.collection('settings').doc('metaOverrides').set(metaOverrides).catch(e => { });
  }

  var key = curLine + '_' + curPart + '_' + curDate;
  if (!saved[key]) saved[key] = { rows: [] };
  localStorage.setItem('seah_insp', JSON.stringify(saved));
  db.collection('inspections').doc(key).set(saved[key]);
  render();
}

function saveData() {
  var allItems = getAllItems();
  var key = curLine + '_' + curPart + '_' + curDate;
  var existing = saved[key] || {};
  var existingRows = existing.rows || [];
  while (existingRows.length < allItems.length) existingRows.push({});

  // GATHER OBSERVATIONS
  var observations = [];
  var norm = function (s) { return String(s || "").replace(/\s+/g, ""); };

  if (curLine === 'OBS') {
    var otrs_obs = document.querySelectorAll('#obody tr');
    for (var k_idx = 0; k_idx < otrs_obs.length; k_idx++) {
      var inps = otrs_obs[k_idx].querySelectorAll('.obs-inp');
      if (inps.length >= 4) {
        var apStr = otrs_obs[k_idx].getAttribute('data-action-photos') || '';
        var ap = apStr ? apStr.split('|') : [];
        observations.push({
          location: (inps[0].value || "").trim(),
          equipment: (inps[1].value || "").trim(),
          detail: (inps[2].value || "").trim(),
          confirm: (inps[3].value || "").trim(),
          actionPhotos: ap
        });
      }
    }
  } else {
    // If not in OBS mode, start with EXISTING observations but STRIP ALL auto-generated ones.
    // They will be re-added accurately by the auto-transfer loop below based on current checklist values.
    // This cleans up any historical contamination or duplicates.
    observations = ((saved[key] || {}).observations || []).filter(function (o) {
      return !(o.detail || "").startsWith("[자동]");
    });
  }

  // Save edits to location/equipment/frequency
  var trs = document.querySelectorAll('#tbody tr');
  var base = getBaseItems(curLine, curPart);
  var ckey = curLine + '_' + curPart;

  for (var i = 0; i < trs.length; i++) {
    var idx = parseInt(trs[i].getAttribute('data-idx'));
    if (isNaN(idx)) continue;
    
    var critInp = trs[i].querySelector('input[data-field="criteria"]');
    var weekInp = trs[i].querySelector('input[data-field="weekLabel"]');
    var locInp = trs[i].querySelector('[data-field="location"]');
    var eqInp = trs[i].querySelector('[data-field="equipment"]');
    var freqSel = trs[i].querySelector('[data-field="frequency"]');

    if (locInp && eqInp && freqSel) {
      var locVal = locInp.value.trim();
      var eqVal = eqInp.value.trim();
      var fVal = freqSel.value;

      if (idx >= base.length) {
        var cIdx = idx - base.length;
        if (!customItems[ckey]) customItems[ckey] = [];
        if (customItems[ckey][cIdx]) {
          customItems[ckey][cIdx].location = locVal;
          customItems[ckey][cIdx].equipment = eqVal;
          customItems[ckey][cIdx].frequency = fVal;
          customItems[ckey][cIdx].criteria = critInp ? critInp.value.trim() : (customItems[ckey][cIdx].criteria || "");
        }
      } else {
        if (!metaOverrides[ckey]) metaOverrides[ckey] = {};
        var critVal = critInp ? critInp.value.trim() : "";
        metaOverrides[ckey][idx] = { location: locVal, equipment: eqVal, frequency: fVal, criteria: critVal };
      }
    }

    var dayInps = trs[i].querySelectorAll('.day-col .inp');
    var remInp = trs[i].querySelector('.rem');

    if (dayInps.length >= 7) {
      var photos = (existingRows[idx] || {}).photos || {};
      var critValActual = critInp ? critInp.value.trim() : (existingRows[idx].criteria || "").trim();

      // Update existingRows for current document
      existingRows[idx] = {
        mon: dayInps[0].value, tue: dayInps[1].value, wed: dayInps[2].value, thu: dayInps[3].value,
        fri: dayInps[4].value, sat: dayInps[5].value, sun: dayInps[6].value,
        remarks: remInp ? remInp.value : '',
        criteria: critValActual,
        weekLabel: weekInp ? weekInp.value : '',
        photos: photos
      };

      // Auto-Transfer Anomaly to Observation List
      if (critValActual && critValActual !== "" && critValActual !== "-") {
        if (!window._weekDatesCache || window._weekDatesCur !== curDate) {
          window._weekDatesCache = getWeekDates(curDate);
          window._weekDatesCur = curDate;
        }
        var weekDates = window._weekDatesCache;

        for (var d = 0; d < 7; d++) {
          var val = dayInps[d].value.trim();
          var tKey = curLine + '_' + curPart + '_' + weekDates[d];
          var lText = allItems[idx] ? (allItems[idx].location || "") : (locInp ? locInp.value.trim() : "");
          var eText = allItems[idx] ? allItems[idx].equipment : (eqInp ? eqInp.value.trim() : "");
          var detText = '[자동] 점검 기준(' + critValActual + ') 위배 - 결과: ' + val;

          if (val !== "" && val !== "-" && val !== critValActual && val !== "양호" && val !== "ㅇㅎ") {
            // Part 1: Add/Update Anomaly
            if (!saved[tKey]) saved[tKey] = { rows: [], observations: [] };
            if (!saved[tKey].observations) saved[tKey].observations = [];

            var targetObsList = (tKey === key) ? observations : saved[tKey].observations;
            var found = targetObsList.find(function (o) { return norm(o.location) === norm(lText) && norm(o.equipment) === norm(eText); });
            if (found) {
              if ((found.detail || "").startsWith("[자동]")) found.detail = detText;
            } else {
              targetObsList.push({ location: lText, equipment: eText, detail: detText, confirm: '' });
            }
          } else if (val === "양호" || val === "ㅇㅎ" || val === critValActual) {
            // Part 2: Remove Anomaly
            var partsToClean = [curLine + '_mechanical_' + weekDates[d], curLine + '_electrical_' + weekDates[d]];
            partsToClean.forEach(function (pk) {
              var targetObsList = (pk === key) ? observations : (saved[pk] && saved[pk].observations ? saved[pk].observations : null);
              if (targetObsList) {
                var filtered = targetObsList.filter(function (o) {
                  var isMatch = norm(o.location) === norm(lText) && norm(o.equipment) === norm(eText);
                  return !(isMatch && (o.detail || "").startsWith("[자동]"));
                });
                if (pk === key) observations = filtered;
                else if (saved[pk]) saved[pk].observations = filtered;
              }
            });
          }
        }
      }
    }
  }

  // Update current in-memory state before cloud sync
  saved[key] = { rows: existingRows, observations: observations };

  // Cloud backup and consistency propagation across the week
  if (curLine !== 'OBS' && curLine !== 'DASH') {
      var wDates = getWeekDates(curDate);
      wDates.forEach(function(wd) {
        var pK = curLine + '_' + curPart + '_' + wd;
        if (saved[pK]) {
          saved[pK].rows = JSON.parse(JSON.stringify(existingRows));
          db.collection('inspections').doc(pK).set(saved[pK]);
        }
      });
  }

  localStorage.setItem('seah_insp', JSON.stringify(saved));
  localStorage.setItem('seah_custom', JSON.stringify(customItems));
  localStorage.setItem('seah_meta', JSON.stringify(metaOverrides));

  db.collection('inspections').doc(key).set(saved[key]);
  db.collection('settings').doc('customItems').set(customItems).catch(e => { });
  db.collection('settings').doc('metaOverrides').set(metaOverrides).catch(e => { });
  
  // Also save observations if in OBS mode
  if (curLine === 'OBS') {
    var otrs_o = document.querySelectorAll('#obody tr[data-line]');
    var lineMap = {};
    for (var i_o = 0; i_o < otrs_o.length; i_o++) {
      var l = otrs_o[i_o].getAttribute('data-line'), p = otrs_o[i_o].getAttribute('data-part');
      var d_o = otrs_o[i_o].getAttribute('data-date') || curDate;
      var subK = l + '_' + p + '_' + d_o;
      if (!lineMap[subK]) lineMap[subK] = [];
      var inps = otrs_o[i_o].querySelectorAll('.obs-inp');
      if (inps.length >= 4) {
        var apStr = otrs_o[i_o].getAttribute('data-action-photos') || '';
        var ap = apStr ? apStr.split('|') : [];
        var entry = {
          location: inps[0].value.trim(),
          equipment: inps[1].value.trim(),
          detail: inps[2].value.trim(),
          confirm: inps[3].value.trim(),
          actionPhotos: ap
        };
        if (!entry.equipment && !entry.detail) continue;
        var isDup = lineMap[subK].some(function (o) {
          return norm(o.location) === norm(entry.location) &&
            norm(o.equipment) === norm(entry.equipment) &&
            norm(o.detail) === norm(entry.detail) &&
            norm(o.confirm) === norm(entry.confirm);
        });
        if (!isDup) lineMap[subK].push(entry);
      }
    }
    for (var sk in lineMap) {
      if (!saved[sk]) saved[sk] = { rows: [] };
      saved[sk].observations = lineMap[sk];
      db.collection('inspections').doc(sk).set(saved[sk]);
    }
    localStorage.setItem('seah_insp', JSON.stringify(saved));
  }
  
  editing = false;
  savedBackup = JSON.stringify(saved);
  updateEditUI(); render();
  alert('저장 완료!');
}

function exportExcel() {
  if (curLine === 'DASH') { alert('종합 현황은 엑셀 내보내기를 지원하지 않습니다.'); return; }
  if (typeof XLSX === 'undefined') { alert('엑셀 라이브러리가 로드되지 않았습니다. 잠시 후 다시 시도해주세요.'); return; }

  var allItems = getAllItems();
  var items = allItems;
  if (curLoc !== 'ALL') items = items.filter(function (it) { return it.location === curLoc; });
  if (curFreq !== 'ALL') items = items.filter(function (it) { return (it.frequency || 'D') === curFreq; });

  var key = curLine + '_' + curPart + '_' + curDate;
  var sd = saved[key] || { rows: [], observations: [] };
  var freqLabel = { 'D': '일', 'W': '주', 'M': '월' };

  var aoa = [];
  var fileName = '';
  var pn = curPart === 'mechanical' ? '기계' : '전기';

  if (curLine === 'OBS') {
    aoa.push(['날짜', '라인', '설비명', '이상 내용 및 조치 사항', '확인']);
    var filterLines = (curObsLine === 'ALL') ? ['CPL', 'CRM', 'CGL', '1CCL', '2CCL', '3CCL', 'SSCL'] : [curObsLine];
    var keys = Object.keys(saved).sort(function (a, b) { return b.split('_')[2].localeCompare(a.split('_')[2]); });
    var curY = curDate.split('-')[0], curMY = curDate.substring(0, 7);

    for (var ki = 0; ki < keys.length; ki++) {
      var k = keys[ki];
      var parts = k.split('_');
      if (parts.length < 3) continue;
      var line = parts[0], date = parts[2];
      if (filterLines.indexOf(line) === -1) continue;
      var match = (curPeriod === 'D' && date === curDate) || (curPeriod === 'W' && isSameWeek(date, curDate)) || (curPeriod === 'M' && date.startsWith(curMY)) || (curPeriod === 'Y' && date.startsWith(curY)) || (curPeriod === 'A');
      if (match) {
        var oarr = (saved[k] || {}).observations || [];
        for (var i = 0; i < oarr.length; i++) {
          if (!oarr[i].equipment && !oarr[i].detail) continue;
          aoa.push([date, line, oarr[i].equipment || '', oarr[i].detail || '', oarr[i].confirm || '']);
        }
      }
    }
    fileName = '이상발견통합관리_' + curDate + '.xlsx';
  } else {
    aoa.push(['순번', '위치', '설비명', '주기', '기준', '월', '화', '수', '목', '금', '토', '일', '비고']);
    for (var i = 0; i < items.length; i++) {
      var it = items[i], origIdx = allItems.indexOf(it), r = (sd.rows || [])[origIdx] || {};
      aoa.push([
        i + 1,
        it.location,
        it.equipment,
        freqLabel[it.frequency || 'D'] || '일',
        it.criteria || '',
        r.mon || '',
        r.tue || '',
        r.wed || '',
        r.thu || '',
        r.fri || '',
        r.sat || '',
        r.sun || '',
        r.remarks || ''
      ]);
    }

    fileName = curLine + '_' + pn + '_점검시스템_' + curDate + '.xlsx';
  }

  var ws = XLSX.utils.aoa_to_sheet(aoa);

  // 열 너비 조절
  if (curLine !== 'OBS') {
    ws['!cols'] = [
      { wch: 6 }, { wch: 15 }, { wch: 40 }, { wch: 8 }, { wch: 8 },
      { wch: 8 }, { wch: 8 }, { wch: 8 }, { wch: 8 }, { wch: 8 },
      { wch: 8 }, { wch: 8 }, { wch: 20 }
    ];
  } else {
    ws['!cols'] = [
      { wch: 15 }, { wch: 10 }, { wch: 30 }, { wch: 50 }, { wch: 15 }
    ];
  }

  var wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Sheet1");
  XLSX.writeFile(wb, fileName);
}

function importExcel(e) {
  var file = e.target.files[0];
  if (!file) return;
  var reader = new FileReader();
  reader.onload = function (e) {
    var data = new Uint8Array(e.target.result);
    var workbook = XLSX.read(data, { type: 'array' });
    var firstSheet = workbook.Sheets[workbook.SheetNames[0]];
    var rows = XLSX.utils.sheet_to_json(firstSheet, { header: 1 });

    if (rows.length < 2) { alert("유효한 데이터가 없습니다."); return; }

    var header = rows[0];
    var isObs = header.includes('이상 내용 및 조치 사항');
    var key = curLine + '_' + curPart + '_' + curDate;
    if (!saved[key]) saved[key] = { rows: [], observations: [] };

    if (isObs) {
      var obs = [];
      for (var i = 1; i < rows.length; i++) {
        var r = rows[i];
        if (!r[2] && !r[3]) continue;
        obs.push({ equipment: r[2] || '', detail: r[3] || '', confirm: r[4] || '' });
      }
      saved[key].observations = obs;
    } else {
      var allItems = getAllItems();
      var freqRev = { '일': 'D', '주': 'W', '월': 'M', 'D': 'D', 'W': 'W', 'M': 'M' };
      var hasSeq = (header[0] === '순번' || !isNaN(rows[1][0]));
      var startCol = hasSeq ? 1 : 0;

      for (var i = 1; i < rows.length; i++) {
        var r = rows[i];
        var loc = r[startCol], eq = r[startCol + 1], freq = freqRev[r[startCol + 2]] || 'D', crit = r[startCol + 3] || '';
        var itIdx = -1;
        for (var j = 0; j < allItems.length; j++) {
          if (allItems[j].location === loc && allItems[j].equipment === eq) { itIdx = j; break; }
        }
        if (itIdx === -1) {
          var ckey = curLine + '_' + curPart;
          if (!customItems[ckey]) customItems[ckey] = [];
          customItems[ckey].push({ location: loc, equipment: eq, frequency: freq, criteria: crit });
          itIdx = allItems.length;
          allItems.push(customItems[ckey][customItems[ckey].length - 1]);
        } else {
          // Update criteria if found
          var ckey = curLine + '_' + curPart;
          var base = getBaseItems(curLine, curPart);
          if (itIdx >= base.length) {
            customItems[ckey][itIdx - base.length].criteria = crit;
          } else {
            if (!metaOverrides[ckey]) metaOverrides[ckey] = {};
            metaOverrides[ckey][itIdx] = Object.assign({}, metaOverrides[ckey][itIdx] || base[itIdx], { criteria: crit });
          }
        }

        if (!saved[key].rows[itIdx]) saved[key].rows[itIdx] = {};
        saved[key].rows[itIdx].mon = r[startCol + 4] || '';
        saved[key].rows[itIdx].tue = r[startCol + 5] || '';
        saved[key].rows[itIdx].wed = r[startCol + 6] || '';
        saved[key].rows[itIdx].thu = r[startCol + 7] || '';
        saved[key].rows[itIdx].fri = r[startCol + 8] || '';
        saved[key].rows[itIdx].sat = r[startCol + 9] || '';
        saved[key].rows[itIdx].sun = r[startCol + 10] || '';
        saved[key].rows[itIdx].remarks = r[startCol + 11] || '';
        saved[key].rows[itIdx].criteria = crit;
      }
    }
    alert("데이터를 불러왔습니다. 하단의 [저장하기] 버튼을 눌러야 최종 반영됩니다.");
    render();
  };
  reader.readAsArrayBuffer(file);
}

var curObsKey = ''; var curObsIdx = -1; var isActionPhotoMode = false;

function openActionPhotoModal(line, part, date, idx) {
  curObsKey = line + '_' + part + '_' + date;
  curObsIdx = idx;
  isActionPhotoMode = true;
  curViewerIdx = 0;

  var obs = saved[curObsKey].observations[idx];
  document.getElementById('photoModalTitle').innerText = obs.equipment + ' 조치 사진 관리';
  document.getElementById('photoModal').classList.add('show');
  renderPhotos();
}

function getWeekDates(dateStr) {
  if (!dateStr) return [];
  var d_now = new Date(dateStr + 'T00:00:00');
  var dayShift = d_now.getDay() || 7;
  var monDay = new Date(d_now);
  monDay.setDate(d_now.getDate() - (dayShift - 1));
  var res = [];
  for (var d = 0; d < 7; d++) {
    var dd = new Date(monDay);
    dd.setDate(monDay.getDate() + d);
    var yyyy = dd.getFullYear();
    var mm = String(dd.getMonth() + 1).padStart(2, '0');
    var rr = String(dd.getDate()).padStart(2, '0');
    res.push(yyyy + '-' + mm + '-' + rr);
  }
  return res;
}

function openPhotoModal(key, idx, day, line, part) {
  curPhotoKey = key; curPhotoIdx = idx; curPhotoDay = day; curViewerIdx = 0; isActionPhotoMode = false;
  var dayLabel = { 'mon': '월', 'tue': '화', 'wed': '수', 'thu': '목', 'fri': '금', 'sat': '토', 'sun': '일' }[day];
  var allItems = getAllItems(line, part);
  var item = allItems[idx];
  var title = (item ? item.equipment : '사진 정보 없음') + ' (' + dayLabel + ') 사진 관리';
  document.getElementById('photoModalTitle').innerText = title;
  document.getElementById('photoModal').classList.add('show');
  renderPhotos();
}

function closePhotoModal() {
  document.getElementById('photoModal').classList.remove('show');
  render();
}

function renderPhotos() {
  var photos = [];
  if (isActionPhotoMode) {
    var obs = saved[curObsKey].observations[curObsIdx];
    photos = obs.actionPhotos || [];
  } else {
    var sd = saved[curPhotoKey] || { rows: [] };
    var r = (sd.rows && sd.rows[curPhotoIdx]) ? sd.rows[curPhotoIdx] : {};
    photos = (r.photos && r.photos[curPhotoDay]) ? r.photos[curPhotoDay] : [];
  }

  var stage = document.getElementById('photoStage');
  var counter = document.getElementById('photoCounter');
  var actions = document.getElementById('photoActions');
  var prevBtn = document.getElementById('photoPrev');
  var nextBtn = document.getElementById('photoNext');
  var notice = document.getElementById('photoSaveNotice');

  if (photos.length === 0) {
    stage.innerHTML = '<div style="color:var(--text-muted)">등록된 사진이 없습니다.</div>';
    counter.textContent = '0 / 0';
    prevBtn.style.display = 'none'; nextBtn.style.display = 'none';
    actions.innerHTML = '';
    notice.style.display = 'none';
  } else {
    if (curViewerIdx >= photos.length) curViewerIdx = photos.length - 1;
    if (curViewerIdx < 0) curViewerIdx = 0;

    stage.innerHTML = '<img src="' + photos[curViewerIdx] + '" onclick="window.open(\'' + photos[curViewerIdx] + '\')">';
    counter.textContent = (curViewerIdx + 1) + ' / ' + photos.length;
    prevBtn.style.display = photos.length > 1 ? 'block' : 'none';
    nextBtn.style.display = photos.length > 1 ? 'block' : 'none';

    var delBtn = editing ? '<button class="btn btn-del" style="background:#ef4444; color:white; border:none" onclick="deletePhoto(' + curViewerIdx + ')">현재 사진 삭제</button>' : '';
    actions.innerHTML = delBtn;
    notice.style.display = editing ? 'block' : 'none';
  }

  var uploadArea = document.querySelector('.upload-area');
  if (uploadArea) uploadArea.style.display = editing ? 'block' : 'none';
}

function prevPhoto() {
  var photos = [];
  if (isActionPhotoMode) {
    photos = saved[curObsKey].observations[curObsIdx].actionPhotos || [];
  } else {
    var sd = saved[curPhotoKey] || { rows: [] };
    var r = sd.rows[curPhotoIdx] || {};
    photos = (r.photos || {})[curPhotoDay] || [];
  }
  if (photos.length === 0) return;
  curViewerIdx--;
  if (curViewerIdx < 0) curViewerIdx = photos.length - 1;
  renderPhotos();
}

function nextPhoto() {
  var photos = [];
  if (isActionPhotoMode) {
    photos = saved[curObsKey].observations[curObsIdx].actionPhotos || [];
  } else {
    var sd = saved[curPhotoKey] || { rows: [] };
    var r = sd.rows[curPhotoIdx] || {};
    photos = (r.photos || {})[curPhotoDay] || [];
  }
  if (photos.length === 0) return;
  curViewerIdx++;
  if (curViewerIdx >= photos.length) curViewerIdx = 0;
  renderPhotos();
}

async function uploadPhotos(input) {
  var files = input.files;
  if (!files.length) return;

  var overlay = document.getElementById('uploadOverlay');
  var status = document.getElementById('uploadStatus');
  overlay.classList.add('show');

  for (var i = 0; i < files.length; i++) {
    var file = files[i];
    status.textContent = '업로드 중... (' + (i + 1) + ' / ' + files.length + ')';

    var fileName = Date.now() + '_' + file.name;
    var path = isActionPhotoMode ?
      'action_photos/' + curObsKey + '/' + curObsIdx + '/' + fileName :
      'photos/' + curPhotoKey + '/' + curPhotoIdx + '/' + curPhotoDay + '/' + fileName;

    var ref = storage.ref().child(path);

    try {
      var snapshot = await ref.put(file);
      var url = await snapshot.ref.getDownloadURL();

      if (isActionPhotoMode) {
        var obs = saved[curObsKey].observations[curObsIdx];
        if (!obs.actionPhotos) obs.actionPhotos = [];
        obs.actionPhotos.push(url);
        curViewerIdx = obs.actionPhotos.length - 1;
      } else {
        if (!saved[curPhotoKey]) saved[curPhotoKey] = { rows: [] };
        if (!saved[curPhotoKey].rows) saved[curPhotoKey].rows = [];
        while (saved[curPhotoKey].rows.length <= curPhotoIdx) saved[curPhotoKey].rows.push({});

        if (!saved[curPhotoKey].rows[curPhotoIdx]) saved[curPhotoKey].rows[curPhotoIdx] = {};
        if (!saved[curPhotoKey].rows[curPhotoIdx].photos) saved[curPhotoKey].rows[curPhotoIdx].photos = {};
        if (!saved[curPhotoKey].rows[curPhotoIdx].photos[curPhotoDay]) saved[curPhotoKey].rows[curPhotoIdx].photos[curPhotoDay] = [];

        saved[curPhotoKey].rows[curPhotoIdx].photos[curPhotoDay].push(url);
        curViewerIdx = saved[curPhotoKey].rows[curPhotoIdx].photos[curPhotoDay].length - 1;
      }
    } catch (err) {
      console.error("Upload Error:", err);
      alert("업로드 중 오류가 발생했습니다: " + file.name);
    }
  }

  input.value = '';
  overlay.classList.remove('show');
  renderPhotos();
}

function deletePhoto(idx) {
  if (!confirm('사진을 삭제하시겠습니까?')) return;
  var photos = [];
  if (isActionPhotoMode) {
    photos = saved[curObsKey].observations[curObsIdx].actionPhotos;
  } else {
    photos = saved[curPhotoKey].rows[curPhotoIdx].photos[curPhotoDay];
  }
  photos.splice(idx, 1);

  if (curViewerIdx >= photos.length) curViewerIdx = Math.max(0, photos.length - 1);
  renderPhotos();
}

function checkValueStatus(criteria, value) {
  if (!value || value.trim() === "" || value === "-") return null;
  var cOrigin = (criteria || "").trim();
  var c = cOrigin.replace(/\s+/g, "");
  var vStr = value.replace(/\s+/g, "");

  // 1. 기준과 완전히 동일하게 입력했으면 OK
  if (vStr === c) return 'ok';

  // 2. 정상/양호 동의어 그룹 처리
  var normalGroup = ['양호', '정상', 'ok', 'ㅇㅎ', '없음', '-', ''];
  var isCritNormal = (normalGroup.indexOf(c) >= 0);
  var isValNormal = (normalGroup.indexOf(vStr) >= 0);

  if (isValNormal) {
     return isCritNormal ? 'ok' : 'error';
  }

  // 3. 부정적 징후 처리
  if (vStr === '이상' || vStr === '불량' || vStr === 'x') return 'error';

  // 4. 수치 판정 로직 시작
  var vMatch = vStr.match(/-?\d+\.?\d*/);
  if (!vMatch) return 'error'; // 입력값에 숫자가 없으면 에러
  var vNum = parseFloat(vMatch[0]);

  // 기준에서 모든 숫자 추출
  var pNums = c.match(/-?\d+\.?\d*/g);
  if (pNums) pNums = pNums.map(parseFloat);

  // (1) 공차 판정: A±B
  if (c.indexOf('±') >= 0 && pNums && pNums.length >= 2) {
    var base = pNums[0], tol = pNums[1];
    return (vNum >= base - tol && vNum <= base + tol) ? 'ok' : 'error';
  }

  // (2) 범위 판정: A~B (단위 포함 33bar~44bar 등 처리)
  var hasRangeChar = (c.indexOf('~') >= 0 || (c.indexOf('-') >= 0 && !c.startsWith('-')));
  if (hasRangeChar && pNums) {
    // 44~ 와 같은 열린 범위 처리
    if (pNums.length === 1) {
       if (c.endsWith('~') || c.endsWith('-')) return vNum >= pNums[0] ? 'ok' : 'error';
       if (c.startsWith('~') || c.startsWith('-')) return vNum <= pNums[0] ? 'ok' : 'error';
    }
    // 33~44 와 같은 닫힌 범위 처리
    if (pNums.length >= 2) {
       var min = Math.min(pNums[0], pNums[1]), max = Math.max(pNums[0], pNums[1]);
       return (vNum >= min && vNum <= max) ? 'ok' : 'error';
    }
  }

  // (3) 조건문 판정 (이상, 이하, 초과, 미만)
  if (pNums) {
    var hasOver = c.indexOf('이상') >= 0, hasUnder = c.indexOf('이하') >= 0;
    var hasGt = c.indexOf('초과') >= 0, hasLt = c.indexOf('미만') >= 0;
    
    if (hasOver || hasUnder || hasGt || hasLt) {
      var minVal = Math.min.apply(null, pNums);
      var maxVal = Math.max.apply(null, pNums);
      var ok = true;
      if (hasOver) if (vNum < minVal) ok = false;
      if (hasUnder) if (vNum > maxVal) ok = false;
      if (hasGt) if (vNum <= minVal) ok = false;
      if (hasLt) if (vNum >= maxVal) ok = false;
      return ok ? 'ok' : 'error';
    }
    
    // 단순 고정 수치 일치 확인
    if (pNums.length === 1) {
       return vNum === pNums[0] ? 'ok' : 'error';
    }
  }

  // 6. 모든 조건에 해당하지 않으면 에러
  return 'error';
}

document.addEventListener('DOMContentLoaded', init);
