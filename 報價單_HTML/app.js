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
    templateBuf: null,  // ArrayBuffer；null = 用內建 B5 範本
    batch: []           // [{name, date, note, rows}]：待合併成一份 Word 的報價單
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

  function loadTemplate() {
    chooseFile('.docx', function (buf, filename) {
      // 粗略驗證是不是 docx（zip 檔頭 PK\x03\x04）
      var head = new Uint8Array(buf.slice(0, 4));
      if (head[0] !== 0x50 || head[1] !== 0x4B) throw new Error('這不是有效的 .docx 檔');
      state.templateBuf = buf;
      setPill('pill-template', '自訂範本：' + filename, true);
      showMsg('已套用自訂範本（' + filename + '）');
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
  function templateBytes() {
    return state.templateBuf
      ? new Uint8Array(state.templateBuf)
      : base64ToUint8(TEMPLATE_B5_BASE64);
  }

  // 把一份報價（{name, note, rows}）算好小計並填入範本，回傳 docxtemplater 實例
  function renderQuote(quote) {
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

    var doc = new window.docxtemplater(new PizZip(templateBytes()), {
      paragraphLoop: true,
      linebreaks: true
    });
    doc.render({
      student_name: quote.name,
      courses: courses,
      sum_tuition: money(st), sum_material: money(sm),
      sum_deduction: deduct(sd), sum_total: money(stot),
      note: quote.note
    });
    return doc;
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

  function makeWord() {
    var quote = currentQuote();
    if (!quote) return;
    try {
      saveAs(zipToBlob(renderQuote(quote).getZip()),
             '報價單_' + safeFileName(quote.name) + '.docx');
      showMsg('Word 已下載（B5 版面，可直接用 Word 編輯）');
    } catch (e) {
      showMsg('產生 Word 失敗：' + renderErr(e), 'err');
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

  function mergeQuotes(quotes) {
    var container = null, parts = [], sectPr = '';

    for (var i = 0; i < quotes.length; i++) {
      var zip = renderQuote(quotes[i]).getZip();
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
    $('btn-batch-word').textContent = '📚 下載合併 Word（' + n + ' 份）';
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

  function batchMakeWord() {
    if (!state.batch.length) { showMsg('批次清單是空的，請先加入報價單', 'err'); return; }
    try {
      var zip = mergeQuotes(state.batch);
      var n = state.batch.length;
      var fname = n === 1
        ? '報價單_' + safeFileName(state.batch[0].name) + '.docx'
        : '報價單合併_' + n + '份.docx';
      saveAs(zipToBlob(zip), fname);
      showMsg('已下載合併 Word：共 ' + n + ' 份，每位學生各一頁，可直接整份列印');
    } catch (e) {
      showMsg('合併失敗：' + renderErr(e), 'err');
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
    $('btn-template').addEventListener('click', loadTemplate);
    $('btn-addrow').addEventListener('click', addRow);
    $('btn-word').addEventListener('click', makeWord);
    $('btn-record').addEventListener('click', exportRecord);
    $('btn-clear').addEventListener('click', clearAll);

    $('btn-batch-add').addEventListener('click', batchAddCurrent);
    $('btn-batch-excel').addEventListener('click', batchLoadExcel);
    $('btn-batch-word').addEventListener('click', batchMakeWord);
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
