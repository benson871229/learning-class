/*!
 * app.js ── 報價單產生器（純前端）
 *
 * 安全守則（本檔案刻意遵守，修改時請一併維持）
 *   1. 絕不使用 innerHTML 塞入任何來自 Excel 或使用者的字串 → 一律 textContent／createElement
 *      （Excel 儲存格可能藏 <script>，這是本工具唯一的外來輸入）
 *   2. 絕不使用 eval / new Function
 *   3. 不做任何網路請求（CSP 也已封鎖）
 *   4. 從 Excel 來的物件以 Object.create(null) 建立（見 xlsx-lite.js），不會汙染原型
 *   5. 匯出 Excel 時中和公式注入（見 xlsx-lite.sanitizeForExcel）
 *
 * 沒有伺服器 → 不存在 LFI／路徑穿越／SSRF／SQL injection 這類伺服器端漏洞。
 */
(function () {
  'use strict';

  var $ = function (id) { return document.getElementById(id); };

  var state = {
    students: [],       // [{name, ...}]
    courses: [],        // [{name, tuition, material, note}]
    tplPayment: null,   // ArrayBuffer；null = 用內建繳費單範本
    tplReceipt: null,   // ArrayBuffer；null = 用內建收據範本
    branch: '壽豐路',    // 收據抬頭要印哪一個分班
    batch: []           // [{name, date, note, rows}]：待合併成一份 Word 的單據
  };

  /* ── 小工具 ── */

  function toNum(v) {
    if (typeof v === 'number') return isFinite(v) ? Math.round(v) : 0;
    var s = String(v == null ? '' : v).replace(/[^0-9.\-]/g, '');
    var n = parseFloat(s);
    return isNaN(n) ? 0 : Math.round(n);
  }

  function fmt(n) { return Number(n || 0).toLocaleString('en-US'); }

  function showMsg(text, kind) {
    var el = $('msg');
    el.textContent = text;                 // textContent：不解析 HTML
    el.className = 'msg ' + (kind || 'ok');
    if (kind !== 'err') {
      window.setTimeout(function () { el.className = 'msg'; }, 4000);
    }
    el.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
  }

  // 從一列資料中，找出第一個名稱包含任一關鍵字的欄位值
  function pick(obj, keywords) {
    var keys = Object.keys(obj);
    for (var i = 0; i < keywords.length; i++) {
      for (var j = 0; j < keys.length; j++) {
        if (keys[j].replace(/\s/g, '').indexOf(keywords[i]) !== -1) return obj[keys[j]];
      }
    }
    return '';
  }

  /* ── 讀取 Excel ── */

  function chooseFile(accept, onLoad) {
    var input = document.createElement('input');
    input.type = 'file';
    input.accept = accept;
    input.addEventListener('change', function () {
      var file = input.files && input.files[0];
      if (!file) return;
      var reader = new FileReader();
      reader.onerror = function () { showMsg('讀取檔案失敗', 'err'); };
      reader.onload = function () {
        try { onLoad(reader.result, file.name); }
        catch (e) { showMsg('無法讀取「' + file.name + '」：' + e.message, 'err'); }
      };
      reader.readAsArrayBuffer(file);
    });
    input.click();
  }

  function loadStudents() {
    chooseFile('.xlsx', function (buf, filename) {
      var objs = XlsxLite.toObjects(XlsxLite.readFirstSheet(buf));
      var list = [];
      for (var i = 0; i < objs.length; i++) {
        var name = String(pick(objs[i], ['學生姓名', '姓名', '學生']) || '').trim();
        if (name) list.push({ name: name, grade: String(pick(objs[i], ['年級']) || '').trim() });
      }
      if (!list.length) throw new Error('找不到「學生姓名」欄位或內容為空');
      state.students = list;
      rebuildStudentSelect();
      setPill('pill-students', list.length + ' 位學生', true);
      showMsg('已載入 ' + list.length + ' 位學生（' + filename + '）');
    });
  }

  function loadCourses() {
    chooseFile('.xlsx', function (buf, filename) {
      var objs = XlsxLite.toObjects(XlsxLite.readFirstSheet(buf));
      var list = [];
      for (var i = 0; i < objs.length; i++) {
        var name = String(pick(objs[i], ['課程名稱', '課程', '班別']) || '').trim();
        if (!name) continue;
        list.push({
          name: name,
          tuition: toNum(pick(objs[i], ['學費', '費用', '金額'])),
          material: toNum(pick(objs[i], ['教材費', '教材'])),
          note: String(pick(objs[i], ['備註', '說明']) || '').trim()
        });
      }
      if (!list.length) throw new Error('找不到「課程名稱」欄位或內容為空');
      state.courses = list;
      refreshCourseSelects();
      setPill('pill-courses', list.length + ' 門課程', true);
      showMsg('已載入 ' + list.length + ' 門課程（' + filename + '）');
    });
  }

  function loadTemplate(kind) {
    var isReceipt = kind === 'receipt';
    chooseFile('.docx', function (buf, filename) {
      // 粗略驗證是不是 docx（zip 檔頭 PK\x03\x04）
      var head = new Uint8Array(buf.slice(0, 4));
      if (head[0] !== 0x50 || head[1] !== 0x4B) throw new Error('這不是有效的 .docx 檔');
      if (isReceipt) state.tplReceipt = buf; else state.tplPayment = buf;
      setPill(isReceipt ? 'pill-tpl-receipt' : 'pill-tpl-payment',
              '自訂範本：' + filename, true);
      showMsg('已套用自訂' + (isReceipt ? '收據' : '繳費單') + '範本（' + filename + '）');
    });
  }

  function setPill(id, text, ok) {
    var el = $(id);
    el.textContent = text;                 // textContent：檔名可能含特殊字元
    el.className = 'pill' + (ok ? ' ok' : '');
  }

  /* ── 學生下拉 ── */

  function rebuildStudentSelect() {
    var sel = $('sel-student');
    while (sel.firstChild) sel.removeChild(sel.firstChild);
    sel.appendChild(new Option('─ 手動輸入 ─', ''));
    for (var i = 0; i < state.students.length; i++) {
      var s = state.students[i];
      var label = s.grade ? s.name + '（' + s.grade + '）' : s.name;
      sel.appendChild(new Option(label, String(i)));   // Option 以文字建立，不解析 HTML
    }
  }

  /* ── 課程列 ── */

  function buildCourseSelect() {
    var sel = document.createElement('select');
    sel.className = 'course-pick';
    fillCourseSelect(sel);
    sel.addEventListener('change', function () { applyCourse(sel); });
    return sel;
  }

  function fillCourseSelect(sel) {
    var keep = sel.value;
    while (sel.firstChild) sel.removeChild(sel.firstChild);
    sel.appendChild(new Option('─ 選課程 ─', ''));
    for (var i = 0; i < state.courses.length; i++) {
      sel.appendChild(new Option(state.courses[i].name, String(i)));
    }
    sel.value = keep;
  }

  function refreshCourseSelects() {
    var sels = document.querySelectorAll('.course-pick');
    for (var i = 0; i < sels.length; i++) fillCourseSelect(sels[i]);
  }

  function numInput(value) {
    var el = document.createElement('input');
    el.type = 'number';
    el.min = '0';
    el.step = '1';
    el.value = String(value == null ? 0 : value);
    el.style.textAlign = 'right';
    el.addEventListener('input', recalc);
    return el;
  }

  function addRow() {
    var tr = document.createElement('tr');

    // 課程名稱：下拉 + 可自由輸入
    var tdName = document.createElement('td');
    var sel = buildCourseSelect();
    sel.style.marginBottom = '4px';
    var nameInput = document.createElement('input');
    nameInput.type = 'text';
    nameInput.placeholder = '課程名稱';
    nameInput.className = 'c-name';
    tdName.appendChild(sel);
    tdName.appendChild(nameInput);

    var tdDate = document.createElement('td');
    var dateInput = document.createElement('input');
    dateInput.type = 'text';
    dateInput.placeholder = '115/6/1-115/6/30';
    dateInput.className = 'c-date';
    tdDate.appendChild(dateInput);

    var tdTuition = document.createElement('td');
    tdTuition.className = 'num';
    var tuition = numInput(0); tuition.className = 'c-tuition';
    tdTuition.appendChild(tuition);

    var tdMaterial = document.createElement('td');
    tdMaterial.className = 'num';
    var material = numInput(0); material.className = 'c-material';
    tdMaterial.appendChild(material);

    var tdDeduct = document.createElement('td');
    tdDeduct.className = 'num';
    var deduct = numInput(0); deduct.className = 'c-deduct';
    tdDeduct.appendChild(deduct);

    var tdTotal = document.createElement('td');
    tdTotal.className = 'num c-total';
    tdTotal.style.fontWeight = '700';
    tdTotal.style.color = '#4f46e5';
    tdTotal.style.paddingTop = '14px';
    tdTotal.textContent = '0';

    var tdDel = document.createElement('td');
    var del = document.createElement('button');
    del.className = 'btn-x';
    del.type = 'button';
    del.textContent = '✕';
    del.addEventListener('click', function () { tr.remove(); recalc(); });
    tdDel.appendChild(del);

    tr.appendChild(tdName); tr.appendChild(tdDate); tr.appendChild(tdTuition);
    tr.appendChild(tdMaterial); tr.appendChild(tdDeduct); tr.appendChild(tdTotal);
    tr.appendChild(tdDel);
    $('rows').appendChild(tr);
    recalc();
  }

  function applyCourse(sel) {
    var idx = parseInt(sel.value, 10);
    if (isNaN(idx) || !state.courses[idx]) return;
    var c = state.courses[idx];
    var tr = sel.closest('tr');
    tr.querySelector('.c-name').value = c.name;
    tr.querySelector('.c-tuition').value = String(c.tuition);
    tr.querySelector('.c-material').value = String(c.material);
    if (c.note && !$('note').value) $('note').value = c.note;
    recalc();
  }

  function readRows() {
    var out = [];
    var trs = $('rows').querySelectorAll('tr');
    for (var i = 0; i < trs.length; i++) {
      var tr = trs[i];
      var name = tr.querySelector('.c-name').value.trim();
      if (!name) continue;
      var tuition = toNum(tr.querySelector('.c-tuition').value);
      var material = toNum(tr.querySelector('.c-material').value);
      var deduction = toNum(tr.querySelector('.c-deduct').value);
      out.push({
        name: name,
        date: tr.querySelector('.c-date').value.trim(),
        tuition: tuition, material: material, deduction: deduction,
        total: tuition + material - deduction
      });
    }
    return out;
  }

  function recalc() {
    var trs = $('rows').querySelectorAll('tr');
    var st = 0, sm = 0, sd = 0, stot = 0;
    for (var i = 0; i < trs.length; i++) {
      var tr = trs[i];
      var tu = toNum(tr.querySelector('.c-tuition').value);
      var ma = toNum(tr.querySelector('.c-material').value);
      var de = toNum(tr.querySelector('.c-deduct').value);
      var tot = tu + ma - de;
      tr.querySelector('.c-total').textContent = fmt(tot);
      st += tu; sm += ma; sd += de; stot += tot;
    }
    $('t-tuition').textContent = fmt(st);
    $('t-material').textContent = fmt(sm);
    $('t-deduct').textContent = sd ? '-' + fmt(sd) : '0';
    $('t-total').textContent = fmt(stot);
  }

  /* ── 產生 Word ── */

  function base64ToUint8(b64) {
    var bin = atob(b64);
    var arr = new Uint8Array(bin.length);
    for (var i = 0; i < bin.length; i++) arr[i] = bin.charCodeAt(i);
    return arr;
  }

  function money(n) { return n === 0 ? '' : String(n); }
  function deduct(n) { return n === 0 ? '' : '-' + n; }

  // 目前的範本內容（自訂優先，否則用內建 B5）
  function templateBytes(kind) {
    if (kind === 'receipt') {
      return state.tplReceipt ? new Uint8Array(state.tplReceipt)
                              : base64ToUint8(TEMPLATE_RECEIPT_BASE64);
    }
    return state.tplPayment ? new Uint8Array(state.tplPayment)
                            : base64ToUint8(TEMPLATE_PAYMENT_BASE64);
  }

  /* ── 分校資料 ──
   * 翠屏路的抬頭刻意沒有「分班」後綴，與壽豐路不同，這是原稿就有的寫法。
   * 證號與電話兩邊相同，屬固定文字寫在範本裡，不在這裡設定。 */
  var BRANCHES = {
    '壽豐路': {
      title: '高雄市私立Fun學院文理短期補習班壽豐分班',
      addr: '高雄市楠梓區壽豐路302號'
    },
    '翠屏路': {
      title: '高雄市私立Fun學院文理短期補習班',
      addr: '高雄市楠梓區翠屏路59號'
    }
  };

  /* 金額轉國字大寫（收據「實收金額」欄用）。
   * 例：12500 → 壹萬貳仟伍佰、100005 → 壹拾萬零伍
   * 組與組之間是否要補「零」，看的是低位那一組有沒有滿千。 */
  function amountToChinese(num) {
    var n = Math.round(Math.abs(Number(num) || 0));
    if (n === 0) return '零';
    var DIGITS = ['零', '壹', '貳', '參', '肆', '伍', '陸', '柒', '捌', '玖'];
    var UNITS = ['', '拾', '佰', '仟'];
    var GROUPS = ['', '萬', '億', '兆'];
    var out = '', gi = 0, prev = null;
    while (n > 0 && gi < GROUPS.length) {
      var group = n % 10000;
      n = Math.floor(n / 10000);
      if (group > 0) {
        var part = '', pendingZero = false, g = group, ui = 0;
        while (g > 0) {
          var d = g % 10;
          g = Math.floor(g / 10);
          if (d === 0) {
            pendingZero = part !== '';
          } else {
            if (pendingZero) { part = '零' + part; pendingZero = false; }
            part = DIGITS[d] + UNITS[ui] + part;
          }
          ui++;
        }
        if (out !== '' && prev !== null && prev < 1000) out = '零' + out;
        out = part + GROUPS[gi] + out;
      }
      prev = group;
      gi++;
    }
    return out;
  }

  // 由課程日期推出收據的修業期間：115/9/1-115/9/30 → 自115年9月1日起至115年9月30日止
  function periodText(quote) {
    var raw = '';
    for (var i = 0; i < quote.rows.length; i++) {
      if (quote.rows[i].date) { raw = quote.rows[i].date; break; }
    }
    if (!raw) return '';
    function ymd(s) {
      var p = String(s).trim().split('/');
      return p.length === 3 ? p[0] + '年' + p[1] + '月' + p[2] + '日' : String(s).trim();
    }
    var parts = raw.split('-');
    return parts.length >= 2
      ? '自' + ymd(parts[0]) + '起至' + ymd(parts[1]) + '止'
      : ymd(raw);
  }

  // 收據內容由繳費單的資料推導：姓名、班別、金額都跟著繳費單走
  function receiptContext(quote) {
    var b = BRANCHES[state.branch] || BRANCHES['壽豐路'];
    var total = 0, names = [];
    for (var i = 0; i < quote.rows.length; i++) {
      total += quote.rows[i].total;
      if (quote.rows[i].name) names.push(quote.rows[i].name);
    }
    return {
      branch_title: b.title,
      branch_addr: b.addr,
      student_name: quote.name,
      class_name: names.join('+'),
      period: periodText(quote),
      amount: String(total),
      amount_cn: amountToChinese(total)
    };
  }

  // 把一份報價（{name, note, rows}）算好小計並填入範本，回傳 docxtemplater 實例

  /* ── 自動縮到一頁 ──
   *
   * 繳費單的版心只有 123mm，基本內容就佔 77mm，每多一門課再多 8.4mm，
   * 所以六門課以上就會被推到第二頁。這裡在套版完成後量一次實際高度，
   * 超出就先拿掉表格裡的空白備用列，再不夠才等比縮小字級與列高。
   *
   * 量測直接讀產生出來的 document.xml，而不是寫死範本的數字，
   * 這樣使用者換上自己的範本一樣有效。
   */

  var TWIP_PER_PT = 20;
  var LINE_FACTOR = 1.30;      // 單行行高相對字級的倍數（中文字型約 1.2~1.4）
  var CELL_PAD = 40;           // 每列框線與內距的餘量（twips）
  var CELL_SIDE_PAD = 216;     // 左右內距 108×2（Word 預設），扣掉才是可排字寬度
  var SAFETY = 0.88;           // 估算約低估一成，留足餘裕免得又爆頁
  var MIN_SZ = 10;             // 字級下限 5pt（w:sz 以半點為單位）
  var MIN_SCALE = 0.35;        // 12 門課又都是超長課名時會用到
  var STYLES = 'word/styles.xml';
  var DOCUMENT = 'word/document.xml';

  function attr(el, name) {
    var v = el && el.getAttribute(name);
    return v == null ? null : parseInt(v, 10);
  }

  function firstSz(el) {
    var list = el.getElementsByTagName('w:sz');
    return list.length ? attr(list[0], 'w:val') : null;
  }

  function maxSz(el) {
    var list = el.getElementsByTagName('w:sz'), m = null;
    for (var i = 0; i < list.length; i++) {
      var v = attr(list[i], 'w:val');
      if (v != null && (m === null || v > m)) m = v;
    }
    return m;
  }

  /* 文字所佔寬度（twips）。中日韓字是全形，數字與英文約半形——
   * 先前一律當全形，收據裡滿是數字（證號、電話、金額、日期），
   * 換行次數因此被嚴重高估，整份高度多算了五成。 */
  function textWidth(s, emTw) {
    var w = 0;
    for (var i = 0; i < s.length; i++) {
      w += s.charCodeAt(i) < 0x2E80 ? emTw * 0.5 : emTw;
    }
    return w;
  }

  function textOf(el) {
    var list = el.getElementsByTagName('w:t'), s = '';
    for (var i = 0; i < list.length; i++) s += list[i].textContent || '';
    return s;
  }

  function parseXml(zip, path) {
    var f = zip.file(path);
    if (!f) return null;
    var d = new DOMParser().parseFromString(f.asText(), 'application/xml');
    return d.getElementsByTagName('parsererror').length ? null : d;
  }

  // 沒有明確 w:sz 的文字都吃 styles.xml 的預設字級，縮放時必須一起處理，
  // 否則只縮到寫死字級的部分，整體高度幾乎不會變。
  function defaultSizeOf(sdoc) {
    if (!sdoc) return 24;
    var dd = sdoc.getElementsByTagName('w:docDefaults')[0];
    return (dd && firstSz(dd)) || 24;
  }

  // 直屬於 el 的 w:spacing（在它自己的 pPr 裡），不抓子元素的
  function ownSpacing(el) {
    var pPr = null, kids = el.childNodes;
    for (var i = 0; i < kids.length; i++) {
      if (kids[i].nodeType === 1 && kids[i].tagName === 'w:pPr') { pPr = kids[i]; break; }
    }
    if (!pPr) return null;
    var k2 = pPr.childNodes;
    for (var j = 0; j < k2.length; j++) {
      if (k2[j].nodeType === 1 && k2[j].tagName === 'w:spacing') return k2[j];
    }
    return null;
  }

  /* 單行行高。除了 w:line（固定值或倍數），還要考慮文件的行網格：
   * sectPr 的 docGrid linePitch 會把每一行撐到固定高度，字級縮小也不會變矮。
   * 這兩份範本的 linePitch 是 360（18pt），收據有 92 個段落受它鎖死，
   * 光網格就吃掉約 584mm——這才是三聯擠不進一頁的主因。 */
  function lineHeight(sp, pt, gridPitch, snaps) {
    var base = pt * TWIP_PER_PT * LINE_FACTOR;
    var rule = sp && sp.getAttribute('w:lineRule');
    var line = sp && attr(sp, 'w:line');
    if (rule === 'exact') return line || base;      // 固定行距不受網格影響
    if (line) base = rule === 'atLeast' ? Math.max(line, base)
                                        : pt * TWIP_PER_PT * (line / 240);
    return snaps && gridPitch ? Math.max(base, gridPitch) : base;
  }

  // 段落是否跟著行網格走（pPr 裡 snapToGrid 設 0 就不跟）
  function snapsToGrid(p) {
    var pPr = null, kids = p.childNodes;
    for (var i = 0; i < kids.length; i++) {
      if (kids[i].nodeType === 1 && kids[i].tagName === 'w:pPr') { pPr = kids[i]; break; }
    }
    if (!pPr) return true;
    var sg = pPr.getElementsByTagName('w:snapToGrid')[0];
    return !(sg && sg.getAttribute('w:val') === '0');
  }

  /* 段落高度。先前漏算了 w:spacing 的 before 與 w:line 倍數，
   * 六門課的繳費單因此被低估約 8mm，才會估得下卻實際爆頁。 */
  function paraHeight(p, defaultSz, widthTw, grid) {
    var pt = (firstSz(p) || defaultSz) / 2;
    var sp = ownSpacing(p);
    var lines = 1;
    if (widthTw > 0) {
      lines = Math.max(1, Math.ceil(textWidth(textOf(p), pt * TWIP_PER_PT) / widthTw));
    }
    var h = lines * lineHeight(sp, pt, grid, snapsToGrid(p));
    if (sp) h += (attr(sp, 'w:before') || 0) + (attr(sp, 'w:after') || 0);
    return h;
  }

  /* 儲存格高度：要把文字換行算進去。
   * 收據的退費規定只宣告 8.6mm 的 trHeight，實際卻要 20mm 以上，
   * 若只信 trHeight 就會嚴重低估整份文件。 */
  function cellHeight(tc, defaultSz, grid) {
    var tcW = tc.getElementsByTagName('w:tcW')[0];
    var w = (attr(tcW, 'w:w') || 0) - CELL_SIDE_PAD;
    var total = 0, ps = tc.getElementsByTagName('w:p');
    for (var i = 0; i < ps.length; i++) total += paraHeight(ps[i], defaultSz, w, grid);
    return total + CELL_PAD;
  }

  function directCells(tr) {
    var out = [], kids = tr.childNodes;
    for (var i = 0; i < kids.length; i++) {
      if (kids[i].nodeType === 1 && kids[i].tagName === 'w:tc') out.push(kids[i]);
    }
    return out;
  }

  function directRows(tbl) {
    var out = [], kids = tbl.childNodes;
    for (var i = 0; i < kids.length; i++) {
      if (kids[i].nodeType === 1 && kids[i].tagName === 'w:tr') out.push(kids[i]);
    }
    return out;
  }

  function vMergeOf(tc) {
    var tcPr = null, kids = tc.childNodes;
    for (var i = 0; i < kids.length; i++) {
      if (kids[i].nodeType === 1 && kids[i].tagName === 'w:tcPr') { tcPr = kids[i]; break; }
    }
    if (!tcPr) return null;
    var vm = tcPr.getElementsByTagName('w:vMerge')[0];
    return vm ? (vm.getAttribute('w:val') || 'continue') : null;
  }

  function declaredHeight(tr) {
    var th = tr.getElementsByTagName('w:trHeight')[0];
    if (!th) return { h: 0, exact: false };
    return { h: attr(th, 'w:val') || 0, exact: th.getAttribute('w:hRule') === 'exact' };
  }

  /* 整個表格各列的高度。
   * 垂直合併（vMerge）的儲存格高度要分攤到它跨越的每一列，不能整份算給起始列——
   * 收據的退費規定就跨了「學費／實收金額／繳費日期」三列，
   * 全算在第一列會讓每一聯多估 12mm，三聯就是 36mm。 */
  function tableRowHeights(tbl, defaultSz, grid) {
    var trs = directRows(tbl), n = trs.length;
    var need = [];
    for (var i = 0; i < n; i++) need.push(0);

    for (var r = 0; r < n; r++) {
      var tcs = directCells(trs[r]);
      for (var c = 0; c < tcs.length; c++) {
        var vm = vMergeOf(tcs[c]);
        if (vm === 'continue') continue;              // 延續列本身不帶內容
        var h = cellHeight(tcs[c], defaultSz, grid);
        if (vm === 'restart') {
          var span = 1;
          for (var k = r + 1; k < n; k++) {
            var kc = directCells(trs[k]);
            if (c < kc.length && vMergeOf(kc[c]) === 'continue') span++;
            else break;
          }
          var per = h / span;
          for (var m = r; m < r + span; m++) need[m] = Math.max(need[m], per);
        } else {
          need[r] = Math.max(need[r], h);
        }
      }
    }

    var total = 0;
    for (var q = 0; q < n; q++) {
      var d = declaredHeight(trs[q]);
      total += d.exact ? d.h : Math.max(d.h, need[q]);
    }
    return total;
  }

  function gridPitchOf(body) {
    var sect = body.getElementsByTagName('w:sectPr')[0];
    var dg = sect && sect.getElementsByTagName('w:docGrid')[0];
    // w:type 的「沒有網格」是 default，不是 none；兩個都要當成沒網格，
    // 否則關掉網格的範本仍會被當成每行鎖 18pt，估高整個偏掉。
    var ty = dg && dg.getAttribute('w:type');
    if (!dg || ty === 'none' || ty === 'default') return 0;
    return attr(dg, 'w:linePitch') || 0;
  }

  function contentHeight(body, defaultSz, widthTw) {
    var grid = gridPitchOf(body);
    var total = 0, kids = body.childNodes;
    for (var i = 0; i < kids.length; i++) {
      var el = kids[i];
      if (el.nodeType !== 1) continue;
      if (el.tagName === 'w:p') {
        total += paraHeight(el, defaultSz, widthTw, grid);
      } else if (el.tagName === 'w:tbl') {
        total += tableRowHeights(el, defaultSz, grid);
      }
    }
    return total;
  }

  function scaleTags(xdoc, tag, k, floor) {
    var list = xdoc.getElementsByTagName(tag);
    for (var i = 0; i < list.length; i++) {
      var v = attr(list[i], 'w:val');
      if (!v) continue;
      // 只縮不放：本來就小於下限的字（例如 6.5pt 的退費規定）若直接套下限，
      // 反而會被放大，整份文件越縮越高。
      var next = Math.max(floor, Math.round(v * k));
      list[i].setAttribute('w:val', String(Math.min(v, next)));
    }
  }

  function scaleAll(xdoc, sdoc, k) {
    var docs = sdoc ? [xdoc, sdoc] : [xdoc];
    for (var i = 0; i < docs.length; i++) {
      scaleTags(docs[i], 'w:sz', k, MIN_SZ);
      scaleTags(docs[i], 'w:szCs', k, MIN_SZ);
    }
    scaleTags(xdoc, 'w:trHeight', k, 1);
    // 行網格：不縮它的話，字縮了行高仍被鎖在 18pt，整份高度幾乎不動
    var dgs = xdoc.getElementsByTagName('w:docGrid');
    for (var g = 0; g < dgs.length; g++) {
      var lp = attr(dgs[g], 'w:linePitch');
      if (lp) dgs[g].setAttribute('w:linePitch', String(Math.max(1, Math.round(lp * k))));
    }
    // 段落間距也要一起收，否則縮了字級、間距原樣佔位
    var sps = xdoc.getElementsByTagName('w:spacing');
    for (var m = 0; m < sps.length; m++) {
      var sp = sps[m];
      ['w:before', 'w:after'].forEach(function (a) {
        var v = attr(sp, a);
        if (v) sp.setAttribute(a, String(Math.round(v * k)));
      });
      if (sp.getAttribute('w:lineRule') === 'exact') {
        var l = attr(sp, 'w:line');
        if (l) sp.setAttribute('w:line', String(Math.round(l * k)));
      }
    }
  }

  // 拿掉表格裡完全空白的列。原稿留著給手寫，但系統是依課程數自動長列，
  // 留著只會佔掉版面，使用者也不想看到。
  function dropBlankRows(body) {
    var tbls = body.getElementsByTagName('w:tbl');
    for (var i = 0; i < tbls.length; i++) {
      var trs = tbls[i].getElementsByTagName('w:tr');
      for (var j = trs.length - 1; j >= 0; j--) {
        if (!textOf(trs[j]).trim()) trs[j].parentNode.removeChild(trs[j]);
      }
    }
  }

  /* 套版後把內容收進一頁：先刪空白列，再等比縮小字級、列高與間距。
   * 高度是估算的，所以採多次收斂；量測直接讀產生出來的 XML，
   * 因此使用者換上自訂範本一樣有效。 */
  /* 頁首／頁尾佔掉的版面。
   * 頁首從 w:header 處開始往下排，比上邊界低就會把本文往下推；
   * 頁尾從 w:footer 處往上排，同理會從下方吃掉空間。
   * 兩邊都要扣，只扣頁首會把可用高度多算十幾 mm，誤判成放得下。 */
  function partHeight(zip, prefix, defaultSz) {
    var pdoc = null;
    for (var i = 1; i <= 3 && !pdoc; i++) pdoc = parseXml(zip, 'word/' + prefix + i + '.xml');
    if (!pdoc) return 0;
    var kids = pdoc.documentElement.childNodes, h = 0;
    for (var j = 0; j < kids.length; j++) {
      if (kids[j].nodeType === 1 && kids[j].tagName === 'w:p') {
        h += paraHeight(kids[j], defaultSz, 0, 0);
      }
    }
    return h;
  }

  function marginOverflow(zip, sect, defaultSz) {
    var pgMar = sect.getElementsByTagName('w:pgMar')[0];
    if (!pgMar) return 0;
    var over = 0;
    if (sect.getElementsByTagName('w:headerReference')[0]) {
      over += Math.max(0, (attr(pgMar, 'w:header') || 0)
                       + partHeight(zip, 'header', defaultSz)
                       - (attr(pgMar, 'w:top') || 0));
    }
    if (sect.getElementsByTagName('w:footerReference')[0]) {
      over += Math.max(0, (attr(pgMar, 'w:footer') || 0)
                       + partHeight(zip, 'footer', defaultSz)
                       - (attr(pgMar, 'w:bottom') || 0));
    }
    return over;
  }

  function fitToOnePage(zip) {
    var xdoc = parseXml(zip, DOCUMENT);
    if (!xdoc) return;
    var sdoc = parseXml(zip, STYLES);

    var body = xdoc.getElementsByTagName('w:body')[0];
    if (!body) return;
    var sect = body.getElementsByTagName('w:sectPr')[0];
    if (!sect) return;
    var pgSz = sect.getElementsByTagName('w:pgSz')[0];
    var pgMar = sect.getElementsByTagName('w:pgMar')[0];
    if (!pgSz || !pgMar) return;

    var dsz0 = defaultSizeOf(sdoc);
    var avail = (attr(pgSz, 'w:h') - (attr(pgMar, 'w:top') || 0)
                 - (attr(pgMar, 'w:bottom') || 0)
                 - marginOverflow(zip, sect, dsz0)) * SAFETY;
    var widthTw = (attr(pgSz, 'w:w') || 0) - (attr(pgMar, 'w:left') || 0)
                  - (attr(pgMar, 'w:right') || 0);
    if (!(avail > 0)) return;

    dropBlankRows(body);                       // 空白列一律不留

    var scale = 1;
    if (contentHeight(body, dsz0, widthTw) > avail) {
      // 二分搜尋「放得下的最大縮放比例」。
      // 先前是逐次依當下高度修正，但縮放與高度不是線性關係，
      // 每輪都會多縮一點，十二門課會一路縮到 6.5pt 卻還剩 24mm 空間。
      var lo = MIN_SCALE, hi = 1, best = MIN_SCALE;
      for (var it = 0; it < 7; it++) {
        var mid = (lo + hi) / 2;
        var trial = xdoc.cloneNode(true);
        scaleAll(trial, null, mid);
        var tb = trial.getElementsByTagName('w:body')[0];
        if (contentHeight(tb, Math.max(MIN_SZ, Math.round(dsz0 * mid)), widthTw) <= avail) {
          best = mid; lo = mid;
        } else {
          hi = mid;
        }
      }
      scaleAll(xdoc, sdoc, best);
      scale = best;
    }

    // 供自動化測試核對估高與實際渲染的落差；只讀不寫，不影響產出。
    window.__fit = {
      avail: avail,
      height: contentHeight(body, Math.max(MIN_SZ, Math.round(dsz0 * scale)), widthTw),
      scale: scale
    };

    var ser = new XMLSerializer();
    zip.file(DOCUMENT, ser.serializeToString(xdoc));
    if (sdoc) zip.file(STYLES, ser.serializeToString(sdoc));
  }

  // kind: 'payment'（繳費單）或 'receipt'（收據）
  function renderQuote(quote, kind) {
    var doc = new window.docxtemplater(new PizZip(templateBytes(kind)), {
      paragraphLoop: true,
      linebreaks: true
    });
    doc.render(kind === 'receipt' ? receiptContext(quote) : paymentContext(quote));
    fitToOnePage(doc.getZip());
    return doc;
  }

  function paymentContext(quote) {
    var st = 0, sm = 0, sd = 0, stot = 0, courses = [];
    for (var i = 0; i < quote.rows.length; i++) {
      var r = quote.rows[i];
      st += r.tuition; sm += r.material; sd += r.deduction; stot += r.total;
      courses.push({
        name: r.name, date: r.date,
        tuition: money(r.tuition), material: money(r.material),
        deduction: deduct(r.deduction), total: money(r.total)
      });
    }
    return {
      student_name: quote.name,
      courses: courses,
      sum_tuition: money(st), sum_material: money(sm),
      sum_deduction: deduct(sd), sum_total: money(stot),
      note: quote.note
    };
  }

  // 從 docxtemplater 錯誤中挖出比較好懂的說明
  function renderErr(e) {
    return e.properties && e.properties.errors && e.properties.errors.length
      ? e.properties.errors[0].properties.explanation : e.message;
  }

  function zipToBlob(zip) {
    return zip.generate({
      type: 'blob',
      mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      compression: 'DEFLATE'   // 不加會產生數百 KB 的未壓縮檔
    });
  }

  var DOC_LABEL = { payment: '繳費單', receipt: '收據' };

  function makeWord(kind) {
    var quote = currentQuote();
    if (!quote) return;
    var label = DOC_LABEL[kind];
    try {
      saveAs(zipToBlob(renderQuote(quote, kind).getZip()),
             label + '_' + safeFileName(quote.name) + '.docx');
      showMsg(label + ' Word 已下載' +
              (kind === 'receipt' ? '（' + state.branch + '分校，A4 三聯）' : '（A5 橫式）'));
    } catch (e) {
      showMsg('產生 ' + label + ' 失敗：' + renderErr(e), 'err');
    }
  }

  // 讀取目前表單成一份報價；資料不全時提示並回傳 null
  function currentQuote() {
    var name = $('student-name').value.trim();
    if (!name) { showMsg('請填寫學生姓名', 'err'); return null; }
    var rows = readRows();
    if (!rows.length) { showMsg('請至少填寫一筆課程（含課程名稱）', 'err'); return null; }
    return {
      name: name,
      date: $('issue-date').value.trim(),
      note: $('note').value.trim(),
      rows: rows
    };
  }

  // 檔名淨化：移除路徑分隔字元與控制字元，避免奇怪的檔名
  function safeFileName(s) {
    return String(s).replace(/[\\/:*?"<>|\x00-\x1F]/g, '_').slice(0, 60) || '報價單';
  }

  /* ── 批次合併 ──
   *
   * 作法：每份報價各自用同一份範本 render 一次，再把每份的 <w:body> 內容接起來，
   * 中間插入分頁符，最後保留一組 sectPr（頁面大小／邊界設定）。
   *
   * 因為所有份數都來自同一個範本，styles.xml / numbering.xml / fontTable 等
   * 其餘零件完全相同，所以直接沿用第一份的 zip 當容器即可，不會有樣式衝突。
   */

  var PAGE_BREAK = '<w:p><w:r><w:br w:type="page"/></w:r></w:p>';
  var BATCH_MAX = 500;          // 合理上限，避免手殘匯入超大檔把瀏覽器卡死

  // 取出 <w:body> 的內容；回傳 {inner, sectPr}
  function splitBody(xml) {
    var m = /<w:body>([\s\S]*)<\/w:body>/.exec(xml);
    if (!m) throw new Error('範本結構異常：找不到 <w:body>');
    var inner = m[1];
    var sectPr = '';
    // 檔尾的 sectPr 是「整份文件」的版面設定，合併時只需保留一組
    var s = /<w:sectPr[\s\S]*<\/w:sectPr>\s*$/.exec(inner);
    if (s) {
      sectPr = s[0];
      inner = inner.slice(0, s.index);
    }
    return { inner: inner, sectPr: sectPr };
  }

  /* 同一份範本複製 N 次會產生重複的元素 id。Word 對重複的書籤 id 尚可容忍，
   * 但圖片（wp:docPr）id 重複會被判定為檔案損毀，因此逐份重新編號。 */
  function renumber(xml, seq) {
    xml = xml.replace(/(<w:bookmarkStart[^>]*\sw:id=")(\d+)(")/g, function (_, a, n, c) {
      return a + (seq * 10000 + Number(n)) + c;
    });
    xml = xml.replace(/(<w:bookmarkEnd[^>]*\sw:id=")(\d+)(")/g, function (_, a, n, c) {
      return a + (seq * 10000 + Number(n)) + c;
    });
    xml = xml.replace(/(<wp:docPr[^>]*\sid=")(\d+)(")/g, function (_, a, n, c) {
      return a + (seq * 10000 + Number(n)) + c;
    });
    return xml;
  }

  function mergeQuotes(quotes, kind) {
    var container = null, parts = [], sectPr = '';

    for (var i = 0; i < quotes.length; i++) {
      var zip = renderQuote(quotes[i], kind).getZip();
      var piece = splitBody(zip.file('word/document.xml').asText());
      if (i === 0) { container = zip; sectPr = piece.sectPr; }
      parts.push(renumber(piece.inner, i));
    }

    var xml = container.file('word/document.xml').asText();
    var body = parts.join(PAGE_BREAK) + sectPr;
    container.file('word/document.xml',
      xml.replace(/<w:body>[\s\S]*<\/w:body>/, '<w:body>' + body + '</w:body>'));
    return container;
  }

  function batchTotal(q) {
    var t = 0;
    for (var i = 0; i < q.rows.length; i++) t += q.rows[i].total;
    return t;
  }

  function renderBatchList() {
    var tb = $('batch-rows');
    while (tb.firstChild) tb.removeChild(tb.firstChild);

    for (var i = 0; i < state.batch.length; i++) {
      (function (idx) {
        var q = state.batch[idx];
        var tr = document.createElement('tr');

        function cell(text, cls) {
          var td = document.createElement('td');
          td.textContent = text;          // textContent：姓名可能來自 Excel
          if (cls) td.className = cls;
          return td;
        }

        tr.appendChild(cell(String(idx + 1)));
        tr.appendChild(cell(q.name));
        tr.appendChild(cell(q.date || '─'));
        tr.appendChild(cell(String(q.rows.length), 'num'));
        tr.appendChild(cell(fmt(batchTotal(q)), 'num'));

        var tdDel = document.createElement('td');
        var del = document.createElement('button');
        del.type = 'button';
        del.className = 'btn-x';
        del.textContent = '✕';
        del.title = '從批次移除';
        del.addEventListener('click', function () {
          state.batch.splice(idx, 1);
          renderBatchList();
        });
        tdDel.appendChild(del);
        tr.appendChild(tdDel);

        tb.appendChild(tr);
      })(i);
    }

    var n = state.batch.length;
    setPill('pill-batch', '批次清單：' + n + ' 份', n > 0);
    $('btn-batch-word').textContent = '📚 合併繳費單（' + n + ' 份）';
    $('btn-batch-receipt').textContent = '🧾 合併收據（' + n + ' 份）';
    $('batch-empty').style.display = n ? 'none' : '';
  }

  function batchAddCurrent() {
    var q = currentQuote();
    if (!q) return;
    if (state.batch.length >= BATCH_MAX) {
      showMsg('批次最多 ' + BATCH_MAX + ' 份', 'err'); return;
    }
    state.batch.push(q);
    renderBatchList();
    showMsg('已加入批次：' + q.name + '（目前 ' + state.batch.length + ' 份）');
  }

  // 報名清單：每列一筆課程，依「學生姓名」分組成多份報價單
  function batchLoadExcel() {
    chooseFile('.xlsx', function (buf, filename) {
      var objs = XlsxLite.toObjects(XlsxLite.readFirstSheet(buf));
      var order = [];
      var byName = Object.create(null);     // null 原型：避免 __proto__ 之類的鍵造成汙染

      for (var i = 0; i < objs.length; i++) {
        var row = objs[i];
        var who = String(pick(row, ['學生姓名', '姓名', '學生']) || '').trim();
        var cname = String(pick(row, ['課程名稱', '課程', '班別']) || '').trim();
        if (!who || !cname) continue;

        var tuition = toNum(pick(row, ['學費', '費用']));
        var material = toNum(pick(row, ['教材費', '教材']));
        var deduction = toNum(pick(row, ['扣除金額', '扣除', '折扣']));

        if (!byName[who]) {
          byName[who] = {
            name: who,
            date: String(pick(row, ['開立日期']) || '').trim(),
            note: String(pick(row, ['備註', '說明']) || '').trim(),
            rows: []
          };
          order.push(who);
        }
        // 備註取該學生第一個非空值
        if (!byName[who].note) {
          byName[who].note = String(pick(row, ['備註', '說明']) || '').trim();
        }
        byName[who].rows.push({
          name: cname,
          date: String(pick(row, ['日期起訖', '日期', '起訖']) || '').trim(),
          tuition: tuition, material: material, deduction: deduction,
          total: tuition + material - deduction
        });
      }

      if (!order.length) throw new Error('找不到「學生姓名」與「課程名稱」欄位，或內容為空');
      if (order.length > BATCH_MAX) {
        throw new Error('一次最多 ' + BATCH_MAX + ' 份，這份清單有 ' + order.length + ' 位學生');
      }

      var list = [];
      for (var k = 0; k < order.length; k++) list.push(byName[order[k]]);
      state.batch = list;
      renderBatchList();
      showMsg('已載入 ' + list.length + ' 份報價單（' + filename + '），可直接下載合併 Word');
    });
  }

  function batchMakeWord(kind) {
    if (!state.batch.length) { showMsg('批次清單是空的，請先加入單據', 'err'); return; }
    var label = DOC_LABEL[kind];
    try {
      var zip = mergeQuotes(state.batch, kind);
      var n = state.batch.length;
      var fname = n === 1
        ? label + '_' + safeFileName(state.batch[0].name) + '.docx'
        : label + '合併_' + n + '份.docx';
      saveAs(zipToBlob(zip), fname);
      showMsg('已下載合併' + label + '：共 ' + n + ' 份，每位學生各一頁，可直接整份列印');
    } catch (e) {
      showMsg('合併' + label + '失敗：' + renderErr(e), 'err');
    }
  }

  function batchExportRecord() {
    if (!state.batch.length) { showMsg('批次清單是空的', 'err'); return; }
    var out = [['開立日期', '學生姓名', '課程名稱', '日期起訖', '學費', '教材費', '扣除金額', '總計', '備註']];
    var grand = 0;
    for (var i = 0; i < state.batch.length; i++) {
      var q = state.batch[i];
      for (var j = 0; j < q.rows.length; j++) {
        var r = q.rows[j];
        grand += r.total;
        out.push([q.date, q.name, r.name, r.date, r.tuition, r.material, r.deduction, r.total, q.note]);
      }
    }
    out.push(['', '', '', '', '', '', '總計', grand, '']);
    try {
      // XlsxLite 會中和 = + - @ 開頭的字串，避免 Excel 公式注入
      saveAs(XlsxLite.write(out, '批次報價紀錄'),
             '批次報價紀錄_' + state.batch.length + '份.xlsx');
      showMsg('批次紀錄 Excel 已下載');
    } catch (e) {
      showMsg('匯出失敗：' + e.message, 'err');
    }
  }

  function batchClear() {
    if (!state.batch.length) return;
    state.batch = [];
    renderBatchList();
    showMsg('已清空批次清單');
  }

  /* ── 匯出紀錄 Excel ── */

  function exportRecord() {
    var name = $('student-name').value.trim();
    var rows = readRows();
    if (!name || !rows.length) { showMsg('請先填好學生姓名與課程再匯出', 'err'); return; }

    var date = $('issue-date').value.trim();
    var note = $('note').value.trim();
    var out = [['開立日期', '學生姓名', '課程名稱', '日期起訖', '學費', '教材費', '扣除金額', '總計', '備註']];
    var stot = 0;
    for (var i = 0; i < rows.length; i++) {
      var r = rows[i];
      stot += r.total;
      out.push([date, name, r.name, r.date, r.tuition, r.material, r.deduction, r.total, note]);
    }
    out.push(['', '', '', '小計', '', '', '', stot, '']);

    try {
      // XlsxLite 會自動中和 = + - @ 開頭的字串，避免 Excel 公式注入
      saveAs(XlsxLite.write(out, '報價紀錄'), '報價紀錄_' + safeFileName(name) + '.xlsx');
      showMsg('紀錄 Excel 已下載');
    } catch (e) {
      showMsg('匯出失敗：' + e.message, 'err');
    }
  }

  /* ── 其他 ── */

  function clearAll() {
    $('student-name').value = '';
    $('note').value = '';
    $('sel-student').value = '';
    var tb = $('rows');
    while (tb.firstChild) tb.removeChild(tb.firstChild);
    addRow();
    showMsg('已清空');
  }

  function init() {
    $('btn-students').addEventListener('click', loadStudents);
    $('btn-courses').addEventListener('click', loadCourses);
    $('btn-tpl-payment').addEventListener('click', function () { loadTemplate('payment'); });
    $('btn-tpl-receipt').addEventListener('click', function () { loadTemplate('receipt'); });

    var branchSel = $('sel-branch');
    branchSel.value = state.branch;
    branchSel.addEventListener('change', function () {
      state.branch = this.value;
      showMsg('收據分校已切換為：' + state.branch);
    });
    $('btn-addrow').addEventListener('click', addRow);
    $('btn-word').addEventListener('click', function () { makeWord('payment'); });
    $('btn-receipt').addEventListener('click', function () { makeWord('receipt'); });
    $('btn-record').addEventListener('click', exportRecord);
    $('btn-clear').addEventListener('click', clearAll);

    $('btn-batch-add').addEventListener('click', batchAddCurrent);
    $('btn-batch-excel').addEventListener('click', batchLoadExcel);
    $('btn-batch-word').addEventListener('click', function () { batchMakeWord('payment'); });
    $('btn-batch-receipt').addEventListener('click', function () { batchMakeWord('receipt'); });
    $('btn-batch-record').addEventListener('click', batchExportRecord);
    $('btn-batch-clear').addEventListener('click', batchClear);

    $('sel-student').addEventListener('change', function () {
      var i = parseInt(this.value, 10);
      if (!isNaN(i) && state.students[i]) $('student-name').value = state.students[i].name;
    });

    addRow();
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init);
  } else {
    init();
  }
})();
