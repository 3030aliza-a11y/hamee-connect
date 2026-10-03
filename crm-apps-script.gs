/**
 * HAMEE CRM – Google Apps Script (miễn phí)
 * Lưu toàn bộ dữ liệu CRM của webApp HAMEE vào 1 Google Sheet để mọi máy Admin dùng chung.
 *
 * CÀI ĐẶT (làm 1 lần, khoảng 5 phút):
 * 1. Tạo Google Sheet mới, đặt tên "HAMEE CRM". Không cần tạo tab, script tự tạo.
 * 2. Tiện ích mở rộng → Apps Script → xoá code mẫu → dán toàn bộ file này.
 * 3. Sửa ADMIN_KEY bên dưới thành một chuỗi bí mật (VD: hamee-2026-xyz). Không chia sẻ chuỗi này.
 * 4. Triển khai → Triển khai mới → Loại: Ứng dụng web
 *      Thực thi với tư cách: Tôi
 *      Người có quyền truy cập: Bất kỳ ai
 *    → Triển khai → cấp quyền → copy link Web App (…/exec).
 * 5. Trên webApp: Admin → ☰ → Cài đặt → dán link Web App + ADMIN_KEY → Lưu.
 *
 * Các tab được tạo: members (hội viên), fees (phiếu thu hội phí), care (nhật ký chăm sóc),
 * events (hoạt động / sự kiện, cột regs chứa danh sách tham dự), applications (phiếu đăng ký gia nhập), settings.
 * Ban thư ký có thể mở Sheet để xem / lọc / in. Nên sửa dữ liệu trên webApp để tránh ghi đè.
 */
// Cài theo cách 2 (script.google.com → Dự án mới): điền ID Google Sheet vào đây
const SHEET_ID = "";
const ss_ = () => SHEET_ID ? SpreadsheetApp.openById(SHEET_ID) : SpreadsheetApp.getActive();
const ADMIN_KEY = "DOI-THANH-MA-BI-MAT";
const TABLES = ["members", "fees", "care", "events", "applications", "settings", "site", "news"];

function doGet(e) {
  const p = e.parameter || {};
  // Dữ liệu công khai cho website (ai cũng đọc được): sự kiện (không kèm danh sách người đăng ký), tin tức, nội dung trang
  if (p.action === "public") {
    const ev = readTable("events").map(x => { const r = Array.isArray(x.regs) ? x.regs : []; const o = Object.assign({}, x); delete o.regs; o.regCount = r.length + readTable("public_regs").filter(q => String(q.eventId) === String(x.id) && !r.some(y => y.rid === q.rid)).length; return o; });
    return json({ events: ev, news: readTable("news").filter(n => n.pub !== false), site: readTable("site") });
  }
  if (p.key !== ADMIN_KEY) return json({ error: "sai_ma_khoa" });
  if (p.action === "dump") {
    const out = {};
    TABLES.concat(["public_regs"]).forEach(t => out[t] = readTable(t));
    return json(out);
  }
  return json({ ok: true });
}

function doPost(e) {
  const d = JSON.parse(e.postData.contents);
  // Khách / hội viên đăng ký sự kiện trên website (không cần mã khoá)
  if (d.action === "reg") {
    const sh = sheetOf("public_regs");
    const keys = ["rid", "eventId", "eventTitle", "date", "name", "org", "phone", "email", "mid", "qty", "unit", "amount", "paid", "checked"];
    if (sh.getLastRow() === 0) sh.appendRow(keys);
    sh.appendRow(keys.map(k => { const v = d.reg[k]; return v === undefined || v === null ? "" : (/^0\d+/.test(String(v)) ? "'" + v : v); }));
    return json({ ok: true });
  }
  if (d.key !== ADMIN_KEY) return json({ error: "sai_ma_khoa" });
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    if (d.action === "upload") {
      const ss = ss_(), par = DriveApp.getFileById(ss.getId()).getParents();
      const root = par.hasNext() ? par.next() : DriveApp.getRootFolder();
      const it = root.getFoldersByName("Poster sự kiện HAMEE");
      const folder = it.hasNext() ? it.next() : root.createFolder("Poster sự kiện HAMEE");
      const f = folder.createFile(Utilities.newBlob(Utilities.base64Decode(d.data), d.type || "image/jpeg", d.name || "poster.jpg"));
      try { f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (err) {}
      return json({ ok: true, id: f.getId() });
    }
    if (d.action === "put") {
      Object.keys(d.tables || {}).forEach(t => { if (TABLES.includes(t)) writeTable(t, d.tables[t] || []); });
      backupDaily();
    }
  } finally { lock.releaseLock(); }
  return json({ ok: true });
}

function sheetOf(name) {
  const ss = ss_();
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function readTable(name) {
  const sh = sheetOf(name), v = sh.getDataRange().getValues();
  if (v.length < 2) return [];
  const h = v.shift();
  return v.filter(r => r.some(x => x !== "")).map(r => {
    const o = {};
    h.forEach((k, i) => {
      let x = r[i];
      if (x instanceof Date) x = Utilities.formatDate(x, "Asia/Ho_Chi_Minh", "yyyy-MM-dd");
      if (typeof x === "string" && /^[\[{]/.test(x)) { try { x = JSON.parse(x); } catch (err) {} }
      if (x === "TRUE" || x === true) x = true; else if (x === "FALSE" || x === false) x = false;
      o[k] = x;
    });
    return o;
  });
}

function writeTable(name, rows) {
  const sh = sheetOf(name);
  const keys = [];
  rows.forEach(r => Object.keys(r || {}).forEach(k => { if (!keys.includes(k)) keys.push(k); }));
  sh.clearContents();
  if (!keys.length) return;
  const data = [keys].concat(rows.map(r => keys.map(k => {
    const x = r[k];
    if (x === undefined || x === null) return "";
    if (typeof x === "object") return JSON.stringify(x);
    if (typeof x === "string" && /^\d{4}-\d{2}-\d{2}$/.test(x)) return "'" + x;   // giữ nguyên dạng ngày
    if (typeof x === "string" && /^0\d+/.test(x)) return "'" + x;                 // giữ số 0 đầu SĐT
    return x;
  })));
  sh.getRange(1, 1, data.length, keys.length).setValues(data);
  sh.setFrozenRows(1);
}

// Sao lưu 1 bản mỗi ngày vào thư mục Drive "HAMEE CRM backup"
function backupDaily() {
  const props = PropertiesService.getScriptProperties();
  const today = Utilities.formatDate(new Date(), "Asia/Ho_Chi_Minh", "yyyy-MM-dd");
  if (props.getProperty("lastBackup") === today) return;
  const ss = ss_();
  const it = DriveApp.getFoldersByName("HAMEE CRM backup");
  const folder = it.hasNext() ? it.next() : DriveApp.createFolder("HAMEE CRM backup");
  DriveApp.getFileById(ss.getId()).makeCopy("HAMEE CRM " + today, folder);
  props.setProperty("lastBackup", today);
}

function json(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

// Chạy 1 lần trong trình soạn thảo để cấp quyền (Sheet, Drive) trước khi Triển khai
function capQuyen() {
  ss_().getName();
  DriveApp.getRootFolder();
}
