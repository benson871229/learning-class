/* ── Price table (NT$/hr) ── */
const PRICE_TABLE = {
  '國小': { '個人教學': 600, '小班教學': 350, '大班教學': 250 },
  '國中': { '個人教學': 700, '小班教學': 400, '大班教學': 300 },
  '高中': { '個人教學': 800, '小班教學': 480, '大班教學': 350 },
};

const SUBJECTS = {
  '國小': ['數學', '英語', '國語文', '自然科學', '社會'],
  '國中': ['數學', '英語', '國文', '理化', '生物', '社會', '地理', '歷史'],
  '高中': ['數學', '英語', '國文', '物理', '化學', '生物', '歷史', '地理', '公民'],
};

const DISCOUNT_MAP = { 1: 1, 3: 0.95, 6: 0.90 };

let rowCounter = 0;
let quoteModalInstance = null;

/* ── Helpers ── */
function today() {
  return new Date().toLocaleDateString('zh-TW', { year: 'numeric', month: '2-digit', day: '2-digit' });
}

function futureDate(days) {
  const d = new Date();
  d.setDate(d.getDate() + days);
  return d.toLocaleDateString('zh-TW', { year: 'numeric', month: '2-digit', day: '2-digit' });
}

function generateQuoteNo() {
  const d = new Date();
  const datePart = `${d.getFullYear()}${String(d.getMonth()+1).padStart(2,'0')}${String(d.getDate()).padStart(2,'0')}`;
  const rand = String(Math.floor(Math.random()*900)+100);
  return `QT-${datePart}-${rand}`;
}

function fmt(n) {
  return 'NT$' + Math.round(n).toLocaleString('zh-TW');
}

/* ── Course row ── */
function addCourseRow() {
  rowCounter++;
  const id = rowCounter;
  const tbody = document.getElementById('course-tbody');
  const tr = document.createElement('tr');
  tr.id = `row-${id}`;
  tr.className = 'course-row';
  tr.innerHTML = `
    <td>
      <select onchange="updateSubjects(${id})" id="level-${id}" class="form-select form-select-sm rounded-2">
        <option value="">選擇年級</option>
        <option value="國小">國小</option>
        <option value="國中">國中</option>
        <option value="高中">高中</option>
      </select>
    </td>
    <td>
      <select id="subject-${id}" class="form-select form-select-sm rounded-2" onchange="updatePrice(${id})">
        <option value="">先選年級</option>
      </select>
    </td>
    <td>
      <select id="format-${id}" class="form-select form-select-sm rounded-2" onchange="updatePrice(${id})">
        <option value="">授課形式</option>
        <option value="個人教學">個人教學</option>
        <option value="小班教學">小班教學</option>
        <option value="大班教學">大班教學</option>
      </select>
    </td>
    <td>
      <select id="hrs-${id}" class="form-select form-select-sm rounded-2" onchange="calcRow(${id})">
        <option value="1">1 小時</option>
        <option value="1.5">1.5 小時</option>
        <option value="2">2 小時</option>
      </select>
    </td>
    <td>
      <input type="number" id="perweek-${id}" value="2" min="1" max="7"
        class="form-control form-control-sm text-center rounded-2"
        onchange="calcRow(${id})">
    </td>
    <td>
      <input type="number" id="months-${id}" value="1" min="1" max="24"
        class="form-control form-control-sm text-center rounded-2"
        onchange="calcRow(${id})">
    </td>
    <td class="small text-center" id="unitprice-${id}" style="color:#64748b">─</td>
    <td class="small fw-semibold text-end text-primary" id="subtotal-${id}">─</td>
    <td class="text-center">
      <button onclick="removeRow(${id})" class="btn btn-sm btn-outline-danger rounded-circle px-2 py-1 lh-1" title="刪除">
        <i class="bi bi-x"></i>
      </button>
    </td>`;
  tbody.appendChild(tr);
  updateTotal();
}

function removeRow(id) {
  const row = document.getElementById(`row-${id}`);
  if (row) row.remove();
  updateTotal();
}

function updateSubjects(id) {
  const level = document.getElementById(`level-${id}`).value;
  const subjectSel = document.getElementById(`subject-${id}`);
  subjectSel.innerHTML = '<option value="">選擇科目</option>';
  if (SUBJECTS[level]) {
    SUBJECTS[level].forEach(s => {
      const opt = document.createElement('option');
      opt.value = s;
      opt.textContent = s;
      subjectSel.appendChild(opt);
    });
  }
  updatePrice(id);
}

function updatePrice(id) {
  const level  = document.getElementById(`level-${id}`).value;
  const format = document.getElementById(`format-${id}`).value;
  const cell   = document.getElementById(`unitprice-${id}`);

  if (level && format && PRICE_TABLE[level]?.[format]) {
    cell.textContent = fmt(PRICE_TABLE[level][format]) + '/hr';
  } else {
    cell.textContent = '─';
  }
  calcRow(id);
}

function calcRow(id) {
  const level    = document.getElementById(`level-${id}`).value;
  const format   = document.getElementById(`format-${id}`).value;
  const hrs      = parseFloat(document.getElementById(`hrs-${id}`).value) || 1;
  const perweek  = parseInt(document.getElementById(`perweek-${id}`).value) || 0;
  const months   = parseInt(document.getElementById(`months-${id}`).value) || 0;
  const cell     = document.getElementById(`subtotal-${id}`);

  const unitprice = PRICE_TABLE[level]?.[format] || 0;
  if (!unitprice) { cell.textContent = '─'; updateTotal(); return; }

  const subtotal = unitprice * hrs * perweek * 4 * months;
  cell.textContent = fmt(subtotal);
  updateTotal();
}

function updateTotal() {
  let total = 0;
  document.querySelectorAll('[id^="subtotal-"]').forEach(cell => {
    const val = cell.textContent.replace(/[^0-9]/g, '');
    if (val) total += parseInt(val);
  });

  const months = getMaxMonths();
  let discountRate = 1;
  for (const [m, rate] of Object.entries(DISCOUNT_MAP)) {
    if (months >= parseInt(m)) discountRate = rate;
  }

  const discountAmt = total * (1 - discountRate);
  const finalTotal  = total - discountAmt;

  document.getElementById('summary-subtotal').textContent  = fmt(total);
  document.getElementById('summary-discount').textContent  = discountRate < 1 ? `-${fmt(discountAmt)} (${(1-discountRate)*100}折)` : '─';
  document.getElementById('summary-total').textContent     = fmt(finalTotal);
  document.getElementById('summary-discount-rate').textContent = discountRate < 1 ? `${(discountRate*10).toFixed(0)}折` : '─';
}

function getMaxMonths() {
  let max = 1;
  document.querySelectorAll('[id^="months-"]').forEach(el => {
    const v = parseInt(el.value) || 0;
    if (v > max) max = v;
  });
  return max;
}

/* ── Generate Preview ── */
function generateQuote() {
  const studentName = document.getElementById('student-name').value.trim();
  if (!studentName) { alert('請填寫學生姓名'); return; }

  const rows = document.querySelectorAll('[id^="row-"]');
  if (rows.length === 0) { alert('請至少新增一項課程'); return; }

  let hasValid = false;
  rows.forEach(row => {
    const id = row.id.split('-')[1];
    if (document.getElementById(`subtotal-${id}`).textContent !== '─') hasValid = true;
  });
  if (!hasValid) { alert('請完整填寫課程資訊（年級、科目、授課形式）'); return; }

  buildPreview();

  if (!quoteModalInstance) {
    quoteModalInstance = new bootstrap.Modal(document.getElementById('quoteModal'));
  }
  quoteModalInstance.show();
}

function buildPreview() {
  const studentName = document.getElementById('student-name').value.trim();
  const grade       = document.getElementById('student-grade').value;
  const parentName  = document.getElementById('parent-name').value.trim();
  const phone       = document.getElementById('phone').value.trim();
  const email       = document.getElementById('email').value.trim();
  const address     = document.getElementById('address').value.trim();
  const note        = document.getElementById('note').value.trim();
  const quoteNo     = generateQuoteNo();

  let courseRows = '';
  let subtotalSum = 0;
  document.querySelectorAll('[id^="row-"]').forEach(row => {
    const id      = row.id.split('-')[1];
    const level   = document.getElementById(`level-${id}`).value;
    const subject = document.getElementById(`subject-${id}`).value;
    const format  = document.getElementById(`format-${id}`).value;
    const hrs     = document.getElementById(`hrs-${id}`).value;
    const perweek = document.getElementById(`perweek-${id}`).value;
    const months  = document.getElementById(`months-${id}`).value;
    const up      = document.getElementById(`unitprice-${id}`).textContent;
    const st      = document.getElementById(`subtotal-${id}`).textContent;
    if (st === '─' || !level) return;
    const stNum = parseInt(st.replace(/[^0-9]/g, ''));
    subtotalSum += stNum;
    courseRows += `
      <tr style="border-bottom:1px solid #e5e7eb">
        <td style="padding:8px 12px">${level}${subject}</td>
        <td style="padding:8px 12px;text-align:center">${format}</td>
        <td style="padding:8px 12px;text-align:center">${hrs} 小時</td>
        <td style="padding:8px 12px;text-align:center">${perweek} 堂/週</td>
        <td style="padding:8px 12px;text-align:center">${months} 個月</td>
        <td style="padding:8px 12px;text-align:right">${up}</td>
        <td style="padding:8px 12px;text-align:right;font-weight:600;color:#4f46e5">NT$${stNum.toLocaleString('zh-TW')}</td>
      </tr>`;
  });

  const months    = getMaxMonths();
  let discountRate = 1;
  for (const [m, rate] of Object.entries(DISCOUNT_MAP)) {
    if (months >= parseInt(m)) discountRate = rate;
  }
  const discountAmt = subtotalSum * (1 - discountRate);
  const finalTotal  = subtotalSum - discountAmt;

  const discountRow = discountAmt > 0
    ? `<tr><td colspan="6" style="padding:8px 12px;text-align:right;color:#6b7280">優惠折扣（${(discountRate*10).toFixed(0)}折）</td><td style="padding:8px 12px;text-align:right;color:#ef4444;font-weight:600">-NT$${Math.round(discountAmt).toLocaleString('zh-TW')}</td></tr>`
    : '';

  const noteRow = note ? `<tr><td colspan="7" style="padding:12px;background:#f9fafb;font-size:13px;color:#4b5563"><strong>備註：</strong>${note}</td></tr>` : '';

  document.getElementById('quote-preview').innerHTML = `
    <div style="font-family:'Microsoft JhengHei','Noto Sans TC',sans-serif;font-size:14px;color:#1f2937;line-height:1.6">
      <div style="display:flex;justify-content:space-between;align-items:flex-start;padding-bottom:16px;border-bottom:3px solid #4f46e5;margin-bottom:20px">
        <div>
          <div style="display:flex;align-items:center;gap:8px;margin-bottom:6px">
            <div style="width:36px;height:36px;background:#4f46e5;border-radius:8px;display:flex;align-items:center;justify-content:center;color:white;font-size:18px;font-weight:900">優</div>
            <span style="font-size:22px;font-weight:900;color:#1f2937">優學補習班</span>
          </div>
          <div style="font-size:12px;color:#6b7280;line-height:1.8">
            📍 台北市中山區學習路100號<br>
            📞 (02) 2345-6789 ｜ 📧 info@youxue.edu.tw
          </div>
        </div>
        <div style="text-align:right">
          <div style="font-size:22px;font-weight:900;color:#4f46e5;margin-bottom:6px">課程報價單</div>
          <div style="font-size:12px;color:#6b7280;line-height:1.8">
            報價單編號：${quoteNo}<br>
            開立日期：${today()}<br>
            有效期限：${futureDate(30)}
          </div>
        </div>
      </div>

      <div style="background:#f8fafc;border-radius:10px;padding:14px 18px;margin-bottom:20px">
        <div style="font-weight:700;color:#374151;margin-bottom:10px;font-size:15px">學生 / 客戶資料</div>
        <div style="display:grid;grid-template-columns:1fr 1fr;gap:8px;font-size:13px">
          <div><span style="color:#6b7280">學生姓名：</span><strong>${studentName}</strong></div>
          <div><span style="color:#6b7280">就讀年級：</span><strong>${grade || '─'}</strong></div>
          <div><span style="color:#6b7280">家長姓名：</span><strong>${parentName || '─'}</strong></div>
          <div><span style="color:#6b7280">聯絡電話：</span><strong>${phone || '─'}</strong></div>
          <div><span style="color:#6b7280">Email：</span><strong>${email || '─'}</strong></div>
          <div><span style="color:#6b7280">地址：</span><strong>${address || '─'}</strong></div>
        </div>
      </div>

      <div style="margin-bottom:20px">
        <div style="font-weight:700;color:#374151;margin-bottom:10px;font-size:15px">課程明細</div>
        <div style="overflow-x:auto">
          <table style="width:100%;border-collapse:collapse;font-size:13px">
            <thead>
              <tr style="background:#4f46e5;color:white">
                <th style="padding:10px 12px;text-align:left;font-weight:600">科目</th>
                <th style="padding:10px 12px;text-align:center;font-weight:600">授課形式</th>
                <th style="padding:10px 12px;text-align:center;font-weight:600">每堂時數</th>
                <th style="padding:10px 12px;text-align:center;font-weight:600">每週堂數</th>
                <th style="padding:10px 12px;text-align:center;font-weight:600">月數</th>
                <th style="padding:10px 12px;text-align:right;font-weight:600">單價/hr</th>
                <th style="padding:10px 12px;text-align:right;font-weight:600">小計</th>
              </tr>
            </thead>
            <tbody>
              ${courseRows}
              <tr style="background:#f9fafb">
                <td colspan="6" style="padding:10px 12px;text-align:right;font-weight:600;color:#374151">課程小計</td>
                <td style="padding:10px 12px;text-align:right;font-weight:700">NT$${subtotalSum.toLocaleString('zh-TW')}</td>
              </tr>
              ${discountRow}
              <tr style="background:#eef2ff">
                <td colspan="6" style="padding:12px;text-align:right;font-weight:800;font-size:15px;color:#4f46e5">應付總計</td>
                <td style="padding:12px;text-align:right;font-weight:900;font-size:18px;color:#4f46e5">NT$${Math.round(finalTotal).toLocaleString('zh-TW')}</td>
              </tr>
              ${noteRow}
            </tbody>
          </table>
        </div>
      </div>

      <div style="display:grid;grid-template-columns:1fr 1fr;gap:16px;margin-bottom:20px;font-size:12px">
        <div style="background:#f8fafc;border-radius:8px;padding:12px">
          <div style="font-weight:700;margin-bottom:6px;color:#374151">付款方式</div>
          <div style="color:#6b7280;line-height:1.8">
            ☐ 月繳（每月5日前）<br>
            ☐ 季繳（9.5折優惠）<br>
            ☐ 學期繳（9折優惠）
          </div>
        </div>
        <div style="background:#f8fafc;border-radius:8px;padding:12px">
          <div style="font-weight:700;margin-bottom:6px;color:#374151">注意事項</div>
          <div style="color:#6b7280;line-height:1.8">
            • 報價有效期 30 天<br>
            • 費用以月為單位計算<br>
            • 首月享免費試聽乙堂
          </div>
        </div>
      </div>

      <div style="display:grid;grid-template-columns:1fr 1fr;gap:24px;margin-top:24px;padding-top:16px;border-top:1px solid #e5e7eb;font-size:12px">
        <div>
          <div style="color:#6b7280;margin-bottom:24px">客戶確認簽名</div>
          <div style="border-bottom:1px solid #9ca3af;padding-bottom:2px;margin-bottom:6px"></div>
          <div style="color:#6b7280">簽名：_________________ 日期：_________</div>
        </div>
        <div>
          <div style="color:#6b7280;margin-bottom:24px">補習班代表蓋章</div>
          <div style="border-bottom:1px solid #9ca3af;padding-bottom:2px;margin-bottom:6px"></div>
          <div style="color:#6b7280">簽名：_________________ 日期：_________</div>
        </div>
      </div>

      <div style="margin-top:20px;text-align:center;font-size:11px;color:#9ca3af">
        優學補習班 ─ 台北市中山區學習路100號 ─ (02) 2345-6789 ─ info@youxue.edu.tw
      </div>
    </div>`;
}

function printQuote() {
  window.print();
}

/* ── Init ── */
document.addEventListener('DOMContentLoaded', () => {
  addCourseRow();
});
