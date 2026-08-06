/*!
 * xlsx-lite.js ── 極簡、無外部相依的 Excel (.xlsx) 讀寫模組
 *
 * 為什麼自己寫，而不是用 npm 上的 xlsx 套件？
 *   1. 資安：npm 最新版 xlsx@0.18.5 有 high 等級漏洞且官方標示 "No fix available"
 *            （原型汙染 GHSA-4r6h-8v6p-xvw6、ReDoS GHSA-5pgg-2g8v-p4x9）
 *   2. 體積：該套件約 5.1MB，與「輕量化」目標相反；本模組約 10KB
 *
 * 安全設計（本檔案刻意遵守的規則）
 *   · 不使用 eval / new Function / innerHTML
 *   · 解析結果一律放進 Object.create(null)，並擋掉 __proto__ / constructor /
 *     prototype 等危險鍵名 → 阻斷原型汙染
 *   · 以 DOMParser 解析 XML（瀏覽器不處理外部實體，故無 XXE 風險）
 *   · 對檔案大小、解壓後大小、列數/欄數設上限 → 阻斷 zip bomb 與記憶體耗盡
 *   · 匯出時對以 = + - @ 開頭的字串加上前置單引號 → 阻斷 Excel 公式注入（CSV injection）
 *
 * 需求：PizZip（已隨附於 lib/）
 * 用法：
 *   const rows = XlsxLite.readFirstSheet(arrayBuffer);   // → [[cell,...], ...]
 *   const objs = XlsxLite.toObjects(rows);               // 首列當表頭 → 物件陣列
 *   const blob = XlsxLite.write(rows, '工作表1');        // → Blob(.xlsx)
 */
var XlsxLite = (function () {
  'use strict';

  // ── 上限（防 zip bomb / 記憶體耗盡）──
  var LIMITS = {
    fileBytes: 15 * 1024 * 1024,      // 原始檔 15MB
    entryBytes: 60 * 1024 * 1024,     // 單一解壓後檔案 60MB
    rows: 20000,
    cols: 256,
    sharedStrings: 200000
  };

  // 危險鍵名：避免寫入物件時汙染原型鏈
  var FORBIDDEN_KEYS = ['__proto__', 'constructor', 'prototype'];

  function isForbiddenKey(k) {
    return FORBIDDEN_KEYS.indexOf(String(k)) !== -1;
  }

  function parseXml(text) {
    var doc = new DOMParser().parseFromString(text, 'application/xml');
    var err = doc.getElementsByTagName('parsererror')[0];
    if (err) throw new Error('XML 解析失敗，檔案可能已損毀');
    return doc;
  }

  function entryText(zip, path) {
    var f = zip.file(path);
    if (!f) return null;
    var txt = f.asText();
    if (txt.length > LIMITS.entryBytes) throw new Error('檔案內容過大，已中止');
    return txt;
  }

  // "AB12" → 欄索引 27（0-based）
  function colIndexFromRef(ref) {
    var n = 0;
    for (var i = 0; i < ref.length; i++) {
      var ch = ref.charCodeAt(i);
      if (ch >= 65 && ch <= 90) n = n * 26 + (ch - 64);           // A-Z
      else if (ch >= 97 && ch <= 122) n = n * 26 + (ch - 96);     // a-z
      else break;
    }
    return n - 1;
  }

  // 取得 <t> 的完整文字（含 rich text 的多個 <r><t>）
  function textOf(node) {
    if (!node) return '';
    var ts = node.getElementsByTagName('t');
    if (!ts.length) return '';
    var out = '';
    for (var i = 0; i < ts.length; i++) out += ts[i].textContent || '';
    return out;
  }

  function readSharedStrings(zip) {
    var xml = entryText(zip, 'xl/sharedStrings.xml');
    if (!xml) return [];
    var items = parseXml(xml).getElementsByTagName('si');
    var n = Math.min(items.length, LIMITS.sharedStrings);
    var out = new Array(n);
    for (var i = 0; i < n; i++) out[i] = textOf(items[i]);
    return out;
  }

  // 找出第一張工作表的路徑（優先照 workbook 關聯，找不到則退回 sheet1.xml）
  function firstSheetPath(zip) {
    try {
      var wb = entryText(zip, 'xl/workbook.xml');
      var rels = entryText(zip, 'xl/_rels/workbook.xml.rels');
      if (wb && rels) {
        var sheet = parseXml(wb).getElementsByTagName('sheet')[0];
        var rid = sheet && (sheet.getAttribute('r:id') || sheet.getAttribute('id'));
        if (rid) {
          var rs = parseXml(rels).getElementsByTagName('Relationship');
          for (var i = 0; i < rs.length; i++) {
            if (rs[i].getAttribute('Id') === rid) {
              var t = rs[i].getAttribute('Target') || '';
              t = t.replace(/^\/?xl\//, '').replace(/^\.\//, '');
              // 僅接受工作表目錄下的相對路徑，拒絕任何跳脫（純前端仍防禦性檢查）
              if (t.indexOf('..') !== -1 || t.charAt(0) === '/') break;
              return 'xl/' + t;
            }
          }
        }
      }
    } catch (e) { /* 退回預設路徑 */ }
    return 'xl/worksheets/sheet1.xml';
  }

  /** 讀取第一張工作表 → 二維陣列（字串或數字） */
  function readFirstSheet(arrayBuffer) {
    if (!arrayBuffer) throw new Error('沒有檔案內容');
    if (arrayBuffer.byteLength > LIMITS.fileBytes) {
      throw new Error('檔案超過 ' + (LIMITS.fileBytes / 1024 / 1024) + 'MB，已拒絕讀取');
    }
    if (typeof PizZip === 'undefined') throw new Error('缺少 PizZip（請確認 lib 資料夾完整）');

    var zip = new PizZip(arrayBuffer);
    var xml = entryText(zip, firstSheetPath(zip));
    if (!xml) throw new Error('這不是有效的 Excel 檔（找不到工作表）');

    var shared = readSharedStrings(zip);
    var rowNodes = parseXml(xml).getElementsByTagName('row');
    var rows = [];
    var rowMax = Math.min(rowNodes.length, LIMITS.rows);

    for (var i = 0; i < rowMax; i++) {
      var cells = rowNodes[i].getElementsByTagName('c');
      var row = [];
      for (var j = 0; j < cells.length; j++) {
        var c = cells[j];
        var ref = c.getAttribute('r') || '';
        var ci = ref ? colIndexFromRef(ref) : j;
        if (ci < 0 || ci >= LIMITS.cols) continue;

        var t = c.getAttribute('t');
        var val = '';
        if (t === 's') {
          var vNode = c.getElementsByTagName('v')[0];
          var idx = vNode ? parseInt(vNode.textContent, 10) : -1;
          val = (idx >= 0 && idx < shared.length) ? shared[idx] : '';
        } else if (t === 'inlineStr') {
          val = textOf(c.getElementsByTagName('is')[0]);
        } else if (t === 'b') {
          var b = c.getElementsByTagName('v')[0];
          val = b && b.textContent === '1' ? 'TRUE' : 'FALSE';
        } else {
          var v = c.getElementsByTagName('v')[0];
          var raw = v ? (v.textContent || '') : '';
          if (raw === '') val = '';
          else if (t === 'str') val = raw;
          else {
            var num = Number(raw);
            val = isNaN(num) ? raw : num;
          }
        }
        while (row.length < ci) row.push('');
        row[ci] = val;
      }
      rows.push(row);
    }
    return rows;
  }

  /**
   * 首列當表頭，其餘轉成物件陣列。
   * 物件以 Object.create(null) 建立且過濾危險鍵名 → 不會汙染原型。
   */
  function toObjects(rows) {
    if (!rows || !rows.length) return [];
    var headers = (rows[0] || []).map(function (h) { return String(h == null ? '' : h).trim(); });
    var out = [];
    for (var i = 1; i < rows.length; i++) {
      var row = rows[i] || [];
      var blank = true;
      for (var k = 0; k < row.length; k++) {
        if (String(row[k] == null ? '' : row[k]).trim() !== '') { blank = false; break; }
      }
      if (blank) continue;

      var obj = Object.create(null);
      for (var j = 0; j < headers.length; j++) {
        var key = headers[j];
        if (!key || isForbiddenKey(key)) continue;   // 擋原型汙染
        var cell = row[j];
        obj[key] = cell == null ? '' : cell;
      }
      out.push(obj);
    }
    return out;
  }

  // ── 匯出 ──
  function escapeXml(s) {
    return String(s)
      .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;').replace(/'/g, '&apos;')
      // 移除 XML 1.0 不合法的控制字元
      .replace(/[\x00-\x08\x0B\x0C\x0E-\x1F]/g, '');
  }

  /**
   * 防 Excel 公式注入：以 = + - @ 或 tab/CR 開頭的字串前面加單引號，
   * 讓 Excel 視為純文字而非公式。
   */
  function sanitizeForExcel(v) {
    if (typeof v === 'number') return v;
    var s = String(v == null ? '' : v);
    if (/^[=+\-@\t\r]/.test(s)) return "'" + s;
    return s;
  }

  function colLetter(n) {
    var s = '';
    n = n + 1;
    while (n > 0) {
      var r = (n - 1) % 26;
      s = String.fromCharCode(65 + r) + s;
      n = Math.floor((n - 1) / 26);
    }
    return s;
  }

  /** 二維陣列 → .xlsx Blob（單一工作表，使用 inline strings） */
  function write(rows, sheetName) {
    if (typeof PizZip === 'undefined') throw new Error('缺少 PizZip（請確認 lib 資料夾完整）');
    var name = escapeXml(String(sheetName || '工作表1').slice(0, 31));
    var body = '';

    for (var i = 0; i < rows.length && i < LIMITS.rows; i++) {
      var row = rows[i] || [];
      var cells = '';
      for (var j = 0; j < row.length && j < LIMITS.cols; j++) {
        var ref = colLetter(j) + (i + 1);
        var v = sanitizeForExcel(row[j]);
        if (typeof v === 'number' && isFinite(v)) {
          cells += '<c r="' + ref + '"><v>' + v + '</v></c>';
        } else if (String(v) === '') {
          continue;
        } else {
          cells += '<c r="' + ref + '" t="inlineStr"><is><t xml:space="preserve">' +
                   escapeXml(v) + '</t></is></c>';
        }
      }
      body += '<row r="' + (i + 1) + '">' + cells + '</row>';
    }

    var sheetXml =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
      '<sheetData>' + body + '</sheetData></worksheet>';

    var workbookXml =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ' +
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
      '<sheets><sheet name="' + name + '" sheetId="1" r:id="rId1"/></sheets></workbook>';

    var workbookRels =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/' +
      'relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>';

    var rootRels =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/' +
      'relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>';

    var contentTypes =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-' +
      'officedocument.spreadsheetml.sheet.main+xml"/>' +
      '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-' +
      'officedocument.spreadsheetml.worksheet+xml"/></Types>';

    var zip = new PizZip();
    zip.file('[Content_Types].xml', contentTypes);
    zip.folder('_rels').file('.rels', rootRels);
    zip.folder('xl').file('workbook.xml', workbookXml);
    zip.folder('xl').folder('_rels').file('workbook.xml.rels', workbookRels);
    zip.folder('xl').folder('worksheets').file('sheet1.xml', sheetXml);

    return zip.generate({
      type: 'blob',
      mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      compression: 'DEFLATE'
    });
  }

  return {
    readFirstSheet: readFirstSheet,
    toObjects: toObjects,
    write: write,
    sanitizeForExcel: sanitizeForExcel,
    LIMITS: LIMITS
  };
})();
