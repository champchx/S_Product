/**
 * BSCM DEMO By Champ — ฐานข้อมูลบน Google Sheets + แจ้งเตือนผ่าน LINE
 *
 * วางไฟล์นี้ใน Apps Script ที่ผูกกับ Google Sheet แล้ว Deploy เป็น Web app
 * แต่ละแท็บในชีตเก็บเอกสารหนึ่งชนิด: id | data (JSON) | updatedAt
 *
 * การแจ้งเตือน LINE ตั้งค่าผ่าน Script Properties (Project Settings → Script Properties)
 *   LINE_TOKEN       Channel access token (long-lived) ของ LINE Official Account
 *   LINE_TO          ID ของกลุ่ม/ผู้ใช้ที่จะรับแจ้งเตือน (คั่นด้วยเครื่องหมายจุลภาคได้หลายรายการ)
 *   LINE_SETUP_CODE  รหัสลับสำหรับพิมพ์ "#ลงทะเบียน รหัส" ในกลุ่ม LINE เพื่อให้ระบบจำ ID กลุ่ม
 *   APP_URL          (ไม่บังคับ) ลิงก์หน้าเว็บ แนบท้ายข้อความแจ้งเตือน
 */
const SHEETS = ['riceTypes', 'orders', 'plans', 'settings'];
const APP_TITLE = 'BSCM DEMO';

// เปิด/ปิดการแจ้งเตือนแต่ละประเภท (true = แจ้ง, false = ไม่แจ้ง)
const NOTIFY = {
  planCreated: true,    // ออกใบสั่งผลิต
  planEdited: true,     // แก้ไขใบสั่งผลิต
  planDeleted: true,    // ลบใบสั่งผลิต
  planReordered: true,  // สลับ/แทรกลำดับการผลิต
  orderStarted: true,   // เริ่มการผลิต (กำลังผลิต)
  orderPaused: true,    // หยุดชั่วคราว
  orderResumed: true,   // เริ่มผลิตต่อ
  orderDone: true,      // ผลิตเสร็จ
  orderReopened: true,  // ย้อนกลับเป็นกำลังผลิต
  orderEdited: true,    // แก้ไขรายการผลิต
  orderDeleted: true,   // ลบรายการผลิต
  inputSummary: true    // สรุปเมื่อธุรการใส่น้ำหนัก input
};

/* ============================ ฐานข้อมูล Google Sheets ============================ */

function sheet_(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange('A:B').setNumberFormat('@'); // เก็บเป็นข้อความ กัน Sheets แปลงรหัสเป็นตัวเลข
    sh.appendRow(['id', 'data', 'updatedAt']);
    sh.setFrozenRows(1);
  }
  return sh;
}

// รันครั้งเดียวตอนตั้งค่า เพื่ออนุญาตสิทธิ์และสร้างแท็บให้ครบ
function setup() {
  SHEETS.forEach(sheet_);
}

function readColl_(name) {
  const vals = sheet_(name).getDataRange().getValues();
  const arr = [];
  for (let i = 1; i < vals.length; i++) {
    if (!vals[i][0]) continue;
    try {
      const d = JSON.parse(vals[i][1]);
      d.id = String(vals[i][0]);
      arr.push(d);
    } catch (e) { /* ข้ามแถวที่ JSON เสีย */ }
  }
  return arr;
}

function readAll_() {
  const out = { ok: true };
  SHEETS.forEach(function (name) { out[name] = readColl_(name); });
  return out;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// อ่านข้อมูลทั้งหมด
function doGet() {
  try {
    return json_(readAll_());
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

// ทำงานหนึ่งรายการ: set (เพิ่ม/แก้ไข) หรือ delete (ลบ) และคืนค่าก่อน/หลังไว้ใช้แจ้งเตือน
function applyOp_(b) {
  if (SHEETS.indexOf(b.coll) < 0) throw new Error('bad collection');
  if (!b.id) throw new Error('missing id');
  const sh = sheet_(b.coll);
  const ids = sh.getRange(1, 1, Math.max(sh.getLastRow(), 1), 1).getValues()
    .map(function (r) { return String(r[0]); });
  const row = ids.indexOf(String(b.id)); // 0 = หัวตาราง
  let old = null;
  if (row > 0) {
    try { old = JSON.parse(sh.getRange(row + 1, 2).getValue()); } catch (e) { old = null; }
  }

  if (b.action === 'set') {
    const data = Object.assign({}, b.data);
    delete data.id;
    const json = JSON.stringify(data);
    if (row > 0) {
      sh.getRange(row + 1, 2).setValue(json);
      sh.getRange(row + 1, 3).setValue(new Date());
    } else {
      sh.appendRow([String(b.id), json, new Date()]);
    }
    return { coll: b.coll, id: String(b.id), action: 'set', old: old, nw: data };
  } else if (b.action === 'delete') {
    if (row > 0) sh.deleteRow(row + 1);
    return { coll: b.coll, id: String(b.id), action: 'delete', old: old, nw: null };
  }
  throw new Error('bad action');
}

// รับคำสั่งเดี่ยว หรือ {action:'batch', ops:[...]} หลายรายการในครั้งเดียว
// รับ webhook จาก LINE ด้วย (body ที่มี events)
// ทุกข้อผิดพลาดจะตอบกลับเป็น JSON เสมอ หน้าเว็บจะได้เห็นสาเหตุจริง
function doPost(e) {
  let body;
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    return json_({ ok: false, error: 'bad json' });
  }
  if (body && body.events) return handleLine_(body);

  let lock = null;
  let error = null;
  const changes = [];
  try {
    lock = LockService.getScriptLock();
    lock.waitLock(25000);
    const ops = body.action === 'batch' ? (body.ops || []) : [body];
    ops.forEach(function (o) { changes.push(applyOp_(o)); });
    SpreadsheetApp.flush();
  } catch (err) {
    error = String(err);
  } finally {
    if (lock) {
      try { lock.releaseLock(); } catch (x) { /* ไม่ได้ถือล็อกอยู่ */ }
    }
  }
  if (error) return json_({ ok: false, error: error });

  // แจ้งเตือน LINE หลังบันทึกเสร็จ ถ้าแจ้งไม่ได้ก็ไม่กระทบการบันทึก
  try {
    notifyChanges_(changes, String(body.by || ''));
  } catch (err) {
    console.error('notify failed: ' + err);
  }
  return json_({ ok: true });
}

/* ================================ แจ้งเตือน LINE ================================ */

const TH_MON = ['ม.ค.', 'ก.พ.', 'มี.ค.', 'เม.ย.', 'พ.ค.', 'มิ.ย.', 'ก.ค.', 'ส.ค.', 'ก.ย.', 'ต.ค.', 'พ.ย.', 'ธ.ค.'];

function props_() { return PropertiesService.getScriptProperties(); }

function getTo_() {
  return String(props_().getProperty('LINE_TO') || '').split(',')
    .map(function (s) { return s.trim(); }).filter(Boolean);
}

function fmtDT_(s) {
  const m = /^(\d{4})-(\d{2})-(\d{2})(?:T(\d{2}):(\d{2}))?/.exec(s || '');
  if (!m) return '-';
  let t = (+m[3]) + ' ' + TH_MON[+m[2] - 1] + ' ' + (+m[1] + 543);
  if (m[4]) t += ' ' + m[4] + ':' + m[5];
  return t;
}

function parseMin_(s) {
  const m = /^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2})/.exec(s || '');
  return m ? Date.UTC(+m[1], +m[2] - 1, +m[3], +m[4], +m[5]) / 60000 : NaN;
}

function minsBetween_(a, b) {
  const d = parseMin_(b) - parseMin_(a);
  return d > 0 ? Math.round(d) : 0;
}

function nowLocal_() {
  return Utilities.formatDate(new Date(), 'Asia/Bangkok', "yyyy-MM-dd'T'HH:mm");
}

function hm_(m) {
  const h = Math.floor(m / 60);
  return (h ? h + ' ชม. ' : '') + (m % 60) + ' นาที';
}

function num_(n, d) {
  const k = Math.pow(10, d === undefined ? 2 : d);
  const v = Math.round((+n || 0) * k) / k;
  const parts = String(v).split('.');
  return parts[0].replace(/\B(?=(\d{3})+(?!\d))/g, ',') + (parts[1] ? '.' + parts[1] : '');
}

function pct_(a, t) { return t > 0 ? num_((+a || 0) / t * 100, 1) + '%' : '-'; }

function sumKg_(arr) {
  return (arr || []).reduce(function (a, x) { return a + (+x.kg || 0); }, 0);
}

function stat_(o) {
  return o.status === 'done' ? 'done' : o.status === 'paused' ? 'paused' : 'running';
}

function lossMin_(o) {
  const end = o.status === 'done' ? o.endAt : nowLocal_();
  return (o.pauses || []).reduce(function (a, p) { return a + minsBetween_(p.from, p.to || end); }, 0);
}

function list_(a) { return (a && a.length) ? a.join(', ') : '-'; }

function msg_(title, rows, by) {
  let t = '[' + APP_TITLE + '] ' + title + '\n' + rows.filter(Boolean).join('\n');
  if (by) t += '\nโดย: ' + by;
  const url = props_().getProperty('APP_URL');
  if (url) t += '\n' + url;
  return t.slice(0, 4900);
}

function plabel_(p) {
  return p.cat ? p.cat + (p.reason ? ' – ' + p.reason : '') : (p.reason || 'ไม่ระบุ');
}

function identRows_(o) {
  return [
    'เลขที่: ' + o.no,
    o.docNo ? 'เลขที่เอกสาร: ' + o.docNo : '',
    'ข้าว: ' + o.typeName + (o.code ? ' (' + o.code + ')' : ''),
    'จำนวน: ' + num_(o.qty) + ' กก.'
  ];
}

// สรุปผลการผลิต: เวลา, input, ผลพลอยได้, ข้าวที่ผลิตได้, กำลังผลิต
function summaryRows_(o) {
  const rows = [];
  const done = o.status === 'done';
  const inKg = sumKg_(o.inputs);
  const byKg = sumKg_(o.byproducts);
  let net = 0, total = 0;
  if (done && o.startAt && o.endAt) {
    total = minsBetween_(o.startAt, o.endAt);
    const loss = lossMin_(o);
    net = Math.max(0, total - loss);
    rows.push('เริ่ม: ' + fmtDT_(o.startAt), 'เสร็จ: ' + fmtDT_(o.endAt),
      'เวลาทั้งหมด: ' + hm_(total),
      'เวลาสูญเสียจากการหยุด: ' + hm_(loss) + ' (' + num_(loss / 60) + ' ชม.)',
      'เวลาผลิตจริง: ' + hm_(net));
  }
  if (inKg > 0) {
    rows.push('— น้ำหนัก input —');
    (o.inputs || []).forEach(function (x) {
      if (+x.kg > 0) rows.push('• ' + (x.name || 'ไม่ระบุชื่อ') + (x.tank ? ' [' + x.tank + ']' : '') + ' ' + num_(x.kg) + ' กก. (' + pct_(x.kg, inKg) + ')');
    });
    rows.push('รวม input: ' + num_(inKg) + ' กก. (' + num_(inKg / 1000, 3) + ' ตัน) = 100%');
  }
  if (done && (o.byproducts || []).length) {
    rows.push('— ผลพลอยได้ —');
    o.byproducts.forEach(function (b) {
      rows.push('• ' + b.name + ' ' + num_(b.kg) + ' กก.' + (inKg > 0 ? ' (' + pct_(b.kg, inKg) + ')' : '') +
        (b.jumbo ? ' (' + b.jumbo + ' จัมโบ้' + (b.rem ? ' + เศษ ' + num_(b.rem) : '') + ')' : ''));
    });
    rows.push('รวมผลพลอยได้: ' + num_(byKg) + ' กก.' + (inKg > 0 ? ' (' + pct_(byKg, inKg) + ' ของ input)' : ''));
  }
  if (done && inKg > 0) {
    rows.push('ข้าวที่ผลิตได้ : ' + num_(inKg - byKg) + ' กก. (' + num_((inKg - byKg) / 1000, 3) + ' ตัน) = ' + pct_(inKg - byKg, inKg) + ' ของ input');
    if (net > 0) rows.push('กำลังผลิต: ' + num_(inKg / 1000 / (net / 60)) + ' ตัน/ชม.');
    const outKg = inKg - byKg, ncr = +((o.rep || {}).ncr) || 0;
    const good = outKg > 0 ? Math.max(0, (outKg - ncr) / outKg) : null;
    const avail = total > 0 ? net / total : null;
    const util = net > 0 ? (inKg / 1000 / (net / 60)) / 10 : null;
    if (good != null && avail != null && util != null) {
      rows.push('— OEE —',
        'ค่าของดี: ' + num_(good * 100) + '%' + (ncr ? ' (ข้าวติด NCR ' + num_(ncr) + ' กก.)' : ''),
        'ค่าความพร้อม: ' + num_(avail * 100) + '%',
        'ค่า Utilize: ' + num_(util * 100) + '%',
        'ค่า OEE: ' + num_(good * avail * util * 100) + '%');
    }
  }
  if (done) rows.push('ถังเก็บข้าวที่ผลิตได้: ' + list_(o.outTanks));
  return rows;
}

// คิวแผนการผลิตที่รอเริ่มผลิต เรียงตามลำดับ
function queueText_() {
  const open = readColl_('plans').filter(function (p) { return p.status !== 'started'; });
  open.sort(function (a, b) {
    const sa = a.seq == null ? 1e9 : a.seq, sb = b.seq == null ? 1e9 : b.seq;
    if (sa !== sb) return sa - sb;
    const da = a.plannedDate || '9', db = b.plannedDate || '9';
    if (da !== db) return da < db ? -1 : 1;
    return (a.createdAt || 0) - (b.createdAt || 0);
  });
  if (!open.length) return 'คิวการผลิต: ไม่มีแผนที่รอเริ่มผลิต';
  const lines = open.slice(0, 10).map(function (p, i) {
    return (i + 1) + '. ' + p.no + ' ' + p.typeName + ' ' + num_(p.qty) + ' กก.';
  });
  if (open.length > 10) lines.push('… และอีก ' + (open.length - 10) + ' แผน');
  return 'ลำดับการผลิตปัจจุบัน\n' + lines.join('\n');
}

function diffRows_(old, nw, fields) {
  const rows = [];
  fields.forEach(function (f) {
    const a = old[f[0]], b = nw[f[0]];
    if (JSON.stringify(a == null ? null : a) === JSON.stringify(b == null ? null : b)) return;
    if (f[2] === 'list') rows.push('• แก้ไข' + f[1]);
    else if (f[2] === 'date') rows.push('• ' + f[1] + ': ' + fmtDT_(a) + ' → ' + fmtDT_(b));
    else if (f[2] === 'num') rows.push('• ' + f[1] + ': ' + num_(a) + ' → ' + num_(b) + ' กก.');
    else rows.push('• ' + f[1] + ': ' + (a || '-') + ' → ' + (b || '-'));
  });
  return rows;
}

function planMessages_(changes, by, out) {
  let created = [], edited = [], deleted = [], reverted = [], reordered = false;
  changes.forEach(function (ch) {
    if (ch.coll !== 'plans') return;
    if (ch.action === 'delete') { if (ch.old) deleted.push(ch.old); return; }
    const p = ch.nw, o = ch.old;
    if (!o) { created.push(p); return; }
    if (o.status === 'started' && p.status !== 'started') { reverted.push(p); return; }
    if (p.status === 'started') return; // เริ่มผลิตจากแผน: แจ้งจากรายการผลิตแล้ว
    if (o.status !== 'started' && p.status !== 'started' && o.seq !== p.seq) reordered = true;
    const rows = diffRows_(o, p, [
      ['typeName', 'ชนิดข้าว', 'text'], ['qty', 'จำนวน', 'num'],
      ['plannedDate', 'วันที่ผลิต', 'date'], ['docNo', 'เลขที่เอกสาร', 'text'], ['note', 'หมายเหตุ', 'text']
    ]);
    if (rows.length) edited.push({ p: p, rows: rows });
  });

  let queueUsed = false;
  created.forEach(function (p) {
    if (!NOTIFY.planCreated) return;
    const rows = identRows_(p).concat(['วันที่ผลิต: ' + fmtDT_(p.plannedDate), p.note ? 'หมายเหตุ: ' + p.note : '']);
    if (reordered && NOTIFY.planReordered && !queueUsed) { rows.push(queueText_()); queueUsed = true; }
    out.push(msg_('📋 ออกใบสั่งผลิต', rows, by));
  });
  edited.forEach(function (x) {
    if (NOTIFY.planEdited) out.push(msg_('✏️ แก้ไขใบสั่งผลิต', identRows_(x.p).concat(x.rows), by));
  });
  reverted.forEach(function (p) {
    if (NOTIFY.planEdited) out.push(msg_('↩️ ใบสั่งผลิตกลับเป็นรอเริ่มผลิต', identRows_(p), by));
  });
  deleted.forEach(function (p) {
    if (NOTIFY.planDeleted) out.push(msg_('🗑️ ลบใบสั่งผลิต', ['เลขที่: ' + p.no, 'ข้าว: ' + p.typeName, 'จำนวน: ' + num_(p.qty) + ' กก.'], by));
  });
  if (reordered && NOTIFY.planReordered && !queueUsed) {
    out.push(msg_('🔀 ปรับลำดับการผลิต', [queueText_()], by));
  }
}

function orderMessages_(ch, by, out) {
  if (ch.coll !== 'orders') return;
  const old = ch.old, o = ch.nw;
  if (ch.action === 'delete') {
    if (old && NOTIFY.orderDeleted) {
      out.push(msg_('🗑️ ลบรายการผลิต', ['เลขที่: ' + old.no, 'ข้าว: ' + old.typeName, 'จำนวน: ' + num_(old.qty) + ' กก.'], by));
    }
    return;
  }
  if (!old) {
    if (NOTIFY.orderStarted) {
      out.push(msg_('▶️ เริ่มการผลิต', identRows_(o).concat([
        'เริ่มเมื่อ: ' + fmtDT_(o.startAt),
        'ถัง input: ' + list_(o.inTanks),
        'ถังเก็บข้าวที่ผลิตได้: ' + list_(o.outTanks)
      ]), by));
    }
    return;
  }

  const sOld = stat_(old), sNew = stat_(o);
  const inOld = sumKg_(old.inputs), inNew = sumKg_(o.inputs);
  let emitted = false;

  if (sOld !== sNew) {
    emitted = true;
    if (sNew === 'paused') {
      const ps = o.pauses || [], p = ps[ps.length - 1] || {};
      if (NOTIFY.orderPaused) out.push(msg_('⏸️ หยุดผลิตชั่วคราว', identRows_(o).concat([
        'เหตุผล: ' + plabel_(p), 'หยุดเมื่อ: ' + fmtDT_(p.from)]), by));
    } else if (sNew === 'done') {
      if (NOTIFY.orderDone) out.push(msg_('✅ ผลิตเสร็จ', identRows_(o).concat(summaryRows_(o)), by));
    } else if (sOld === 'paused') {
      const oldPs = old.pauses || [], idx = oldPs.length - 1, np = (o.pauses || [])[idx] || {};
      if (NOTIFY.orderResumed) out.push(msg_('▶️ เริ่มผลิตต่อ', identRows_(o).concat([
        'เริ่มต่อเมื่อ: ' + fmtDT_(np.to), 'หยุดไป: ' + hm_(minsBetween_(np.from, np.to)),
        'เหตุผลที่หยุด: ' + plabel_(np)]), by));
    } else if (sOld === 'done') {
      if (NOTIFY.orderReopened) out.push(msg_('↩️ ย้อนกลับเป็นกำลังผลิต', identRows_(o), by));
    }
  }

  // สรุปเมื่อใส่น้ำหนัก input (ถ้าเพิ่งผลิตเสร็จพร้อมน้ำหนัก ข้อความ "ผลิตเสร็จ" มีสรุปอยู่แล้ว)
  const doneNow = sOld !== sNew && sNew === 'done';
  if (inNew > 0 && Math.round(inNew * 100) !== Math.round(inOld * 100) && !doneNow) {
    emitted = true;
    if (NOTIFY.inputSummary) {
      out.push(msg_(sNew === 'done' ? '⚖️ *สรุปผลการผลิต ' : '⚖️ บันทึกน้ำหนัก input',
        identRows_(o).concat(summaryRows_(o)).concat(sNew === 'done' ? [] : [
          '— น้ำหนัก input —'].concat((o.inputs || []).filter(function (x) { return +x.kg > 0; }).map(function (x) {
          return '• ' + (x.name || 'ไม่ระบุชื่อ') + (x.tank ? ' [' + x.tank + ']' : '') + ' ' + num_(x.kg) + ' กก. (' + pct_(x.kg, inNew) + ')';
        })).concat(['รวม input: ' + num_(inNew) + ' กก. (' + num_(inNew / 1000, 3) + ' ตัน) = 100%'])), by));
    }
  }

  if (!emitted) {
    const rows = diffRows_(old, o, [
      ['typeName', 'ชนิดข้าว', 'text'], ['qty', 'จำนวน', 'num'],
      ['startAt', 'เวลาเริ่มผลิต', 'date'], ['endAt', 'เวลาเสร็จ', 'date'],
      ['due', 'กำหนดเสร็จ', 'date'], ['docNo', 'เลขที่เอกสาร', 'text'], ['note', 'หมายเหตุ', 'text'],
      ['inTanks', 'ถัง input', 'list'], ['outTanks', 'ถังเก็บข้าวที่ผลิตได้', 'list'],
      ['byproducts', 'ผลพลอยได้', 'list'], ['pauses', 'ประวัติการหยุด', 'list'],
      ['inputs', 'น้ำหนัก input', 'list']
    ]);
    if (rows.length && NOTIFY.orderEdited) {
      out.push(msg_('✏️ แก้ไขรายการผลิต', identRows_(o).concat(['สถานะ: ' + ({ running: 'กำลังผลิต', paused: 'หยุดชั่วคราว', done: 'ผลิตเสร็จ' }[sNew])]).concat(rows), by));
    }
  }
}

function notifyChanges_(changes, by) {
  const token = props_().getProperty('LINE_TOKEN');
  const to = getTo_();
  if (!token || !to.length) return;
  const out = [];
  planMessages_(changes, by, out);
  changes.forEach(function (ch) { orderMessages_(ch, by, out); });
  if (out.length) pushLine_(token, to, out);
}

// ส่งหลายข้อความในคำขอเดียว (สูงสุด 5 ข้อความ/คำขอ) เพื่อประหยัดโควตา
function pushLine_(token, to, texts) {
  for (let i = 0; i < texts.length; i += 5) {
    const messages = texts.slice(i, i + 5).map(function (t) { return { type: 'text', text: t }; });
    to.forEach(function (id) {
      try {
        const res = UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', {
          method: 'post',
          contentType: 'application/json',
          headers: { Authorization: 'Bearer ' + token },
          payload: JSON.stringify({ to: id, messages: messages }),
          muteHttpExceptions: true
        });
        if (res.getResponseCode() !== 200) console.error('LINE push ' + res.getResponseCode() + ': ' + res.getContentText());
      } catch (err) {
        console.error('LINE push error: ' + err);
      }
    });
  }
}

function replyLine_(replyToken, text) {
  const token = props_().getProperty('LINE_TOKEN');
  if (!token) { keep_('ไม่พบ LINE_TOKEN ใน Script Properties'); return; }
  if (!replyToken) return;
  const res = UrlFetchApp.fetch('https://api.line.me/v2/bot/message/reply', {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + token },
    payload: JSON.stringify({ replyToken: replyToken, messages: [{ type: 'text', text: text }] }),
    muteHttpExceptions: true
  });
  keep_('LINE reply ' + res.getResponseCode() + ' ' + res.getContentText());
}

// รับ webhook จาก LINE: ใช้ลงทะเบียนกลุ่มที่จะรับแจ้งเตือน
//   #id                    ตอบ ID ของแชตนี้
//   #ลงทะเบียน <รหัส>      เพิ่มแชตนี้เป็นผู้รับแจ้งเตือน (รหัสต้องตรงกับ LINE_SETUP_CODE)
//   #เลิกรับ               เอาแชตนี้ออกจากผู้รับแจ้งเตือน
function handleLine_(body) {
  (body.events || []).forEach(function (ev) {
    try {
      const src = ev.source || {};
      keep_('LINE event: ' + JSON.stringify(ev).slice(0, 400));
      const id = src.groupId || src.roomId || src.userId;
      if (ev.type === 'join' || ev.type === 'follow') {
        replyLine_(ev.replyToken, 'สวัสดีครับ ระบบแจ้งเตือน ' + APP_TITLE + '\nพิมพ์ #ลงทะเบียน ตามด้วยรหัส เพื่อรับแจ้งเตือนในแชตนี้\nพิมพ์ #id เพื่อดู ID ของแชตนี้');
        return;
      }
      if (ev.type !== 'message' || !ev.message || ev.message.type !== 'text') return;
      const text = String(ev.message.text || '').trim();
      const to = getTo_();
      if (/^#id$/i.test(text)) {
        replyLine_(ev.replyToken, 'ID ของแชตนี้:\n' + id);
        return;
      }
      const m = /^#ลงทะเบียน\s*(.*)$/.exec(text);
      if (m) {
        const code = props_().getProperty('LINE_SETUP_CODE');
        if (!code) { replyLine_(ev.replyToken, 'ยังไม่ได้ตั้งค่า LINE_SETUP_CODE ใน Apps Script'); return; }
        if (m[1].trim() !== code) { replyLine_(ev.replyToken, 'รหัสไม่ถูกต้อง'); return; }
        if (to.indexOf(id) < 0) { to.push(id); props_().setProperty('LINE_TO', to.join(',')); }
        replyLine_(ev.replyToken, '✅ ลงทะเบียนแล้ว แชตนี้จะได้รับแจ้งเตือนจากระบบ ' + APP_TITLE);
        return;
      }
      if (text === '#เลิกรับ') {
        const rest = to.filter(function (x) { return x !== id; });
        props_().setProperty('LINE_TO', rest.join(','));
        replyLine_(ev.replyToken, 'ยกเลิกการแจ้งเตือนในแชตนี้แล้ว');
      }
    } catch (err) {
      keep_('line event error: ' + err);
    }
  });
  return json_({ ok: true });
}

// เรียกใช้จากตัวแก้ไข Apps Script เพื่อทดสอบว่าส่งเข้า LINE ได้
function testLine() {
  const token = props_().getProperty('LINE_TOKEN');
  const to = getTo_();
  if (!token) throw new Error('ยังไม่ได้ตั้งค่า LINE_TOKEN');
  if (!to.length) throw new Error('ยังไม่ได้ตั้งค่า LINE_TO (พิมพ์ #ลงทะเบียน รหัส ในกลุ่ม LINE หรือใส่ ID เอง)');
  pushLine_(token, to, ['[' + APP_TITLE + '] ✅ ทดสอบการแจ้งเตือนสำเร็จ']);
}

// เก็บบันทึกล่าสุดไว้ใน Script Properties (เผื่อหน้า Executions ไม่แสดง log)
function keep_(t) {
  console.log(t);
  try {
    const old = String(props_().getProperty('LAST_LOG') || '');
    const line = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'HH:mm:ss') + ' ' + t;
    props_().setProperty('LAST_LOG', (old + '\n' + line).slice(-1800));
  } catch (e) { /* ignore */ }
}

// รันจากตัวแก้ไข Apps Script เพื่อตรวจ Token และดู log ล่าสุดของ webhook
function diag() {
  const token = props_().getProperty('LINE_TOKEN');
  console.log('LINE_TOKEN: ' + (token ? 'มี (ยาว ' + token.length + ' ตัวอักษร)' : 'ไม่มี'));
  console.log('LINE_SETUP_CODE: ' + (props_().getProperty('LINE_SETUP_CODE') ? 'มี' : 'ไม่มี'));
  console.log('LINE_TO: ' + (props_().getProperty('LINE_TO') || '(ว่าง)'));
  if (token) {
    const r = UrlFetchApp.fetch('https://api.line.me/v2/bot/info', {
      headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true });
    console.log('ตรวจ Token กับ LINE: ' + r.getResponseCode() + ' ' + r.getContentText());
  }
  console.log('--- log webhook ล่าสุด ---' + (props_().getProperty('LAST_LOG') || '\n(ยังไม่มี: LINE ยังไม่เคยส่งข้อความมาถึงโค้ดเวอร์ชันนี้)'));
}
