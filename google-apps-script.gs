/**
 * بک‌اند داشبورد ارزیابی استخدام
 * -------------------------------------------------
 * نحوه نصب:
 * 1) یک گوگل‌شیت جدید بسازید (sheet.new)
 * 2) از منو: Extensions > Apps Script
 * 3) محتوای این فایل را جایگزین کد پیش‌فرض کنید و ذخیره کنید (Ctrl+S)
 * 4) از دکمه Deploy > New deployment
 *    - Select type: Web app
 *    - Execute as: Me
 *    - Who has access: Anyone
 * 5) روی Deploy بزنید و لینک "Web app URL" را کپی کنید
 * 6) همان لینک را داخل داشبورد، در کادر "اتصال به گوگل‌شیت" بچسبانید
 *
 * این اسکریپت به‌صورت خودکار شیت‌های Config / Positions / Candidates / Evaluations
 * را در همین spreadsheet می‌سازد و مدیریت می‌کند.
 */

function doGet(e) {
  const data = {
    positions: readPositions(),
    stages: readConfig().stages,
    candidates: readCandidates(),
    evaluations: readEvaluations(),
  };
  return jsonResponse(data);
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    writeConfig({ stages: body.stages || [] });
    writePositions(body.positions || []);
    writeCandidates(body.candidates || []);
    writeEvaluations(body.evaluations || []);
    return jsonResponse({ ok: true });
  } catch (err) {
    return jsonResponse({ ok: false, error: String(err) });
  }
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function ss() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function getOrCreateSheet(name) {
  let sh = ss().getSheetByName(name);
  if (!sh) sh = ss().insertSheet(name);
  return sh;
}

// ---------- Config (stages + criteria + weights) ----------
function writeConfig(config) {
  const sh = getOrCreateSheet("Config");
  sh.clear();
  sh.getRange(1, 1).setValue(JSON.stringify(config));
}
function readConfig() {
  const sh = getOrCreateSheet("Config");
  const raw = sh.getRange(1, 1).getValue();
  if (!raw) return { stages: [] };
  try { return JSON.parse(raw); } catch (e) { return { stages: [] }; }
}

// ---------- Positions ----------
function writePositions(positions) {
  const sh = getOrCreateSheet("Positions");
  sh.clear();
  sh.appendRow(["id", "title"]);
  positions.forEach(p => sh.appendRow([p.id, p.title]));
}
function readPositions() {
  const sh = getOrCreateSheet("Positions");
  const values = sh.getDataRange().getValues();
  const rows = values.slice(1);
  return rows.filter(r => r[0]).map(r => ({ id: r[0], title: r[1] }));
}

// ---------- Candidates ----------
function writeCandidates(candidates) {
  const sh = getOrCreateSheet("Candidates");
  sh.clear();
  sh.appendRow(["id", "name", "positionId", "status", "createdAt"]);
  candidates.forEach(c => sh.appendRow([c.id, c.name, c.positionId, c.status, c.createdAt]));
}
function readCandidates() {
  const sh = getOrCreateSheet("Candidates");
  const values = sh.getDataRange().getValues();
  const rows = values.slice(1);
  return rows.filter(r => r[0]).map(r => ({
    id: r[0], name: r[1], positionId: r[2], status: r[3] || "pending", createdAt: r[4],
  }));
}

// ---------- Evaluations ----------
function writeEvaluations(evaluations) {
  const sh = getOrCreateSheet("Evaluations");
  sh.clear();
  sh.appendRow(["id", "candidateId", "stageId", "evaluatorName", "scoresJSON", "note", "updatedAt"]);
  evaluations.forEach(ev => sh.appendRow([
    ev.id, ev.candidateId, ev.stageId, ev.evaluatorName,
    JSON.stringify(ev.scores || {}), ev.note || "", ev.updatedAt || Date.now(),
  ]));
}
function readEvaluations() {
  const sh = getOrCreateSheet("Evaluations");
  const values = sh.getDataRange().getValues();
  const rows = values.slice(1);
  return rows.filter(r => r[0]).map(r => {
    let scores = {};
    try { scores = JSON.parse(r[4] || "{}"); } catch (e) {}
    return {
      id: r[0], candidateId: r[1], stageId: r[2], evaluatorName: r[3],
      scores, note: r[5] || "", updatedAt: r[6],
    };
  });
}
