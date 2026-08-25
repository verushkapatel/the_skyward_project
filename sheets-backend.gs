/**
 * FinLit Index · Pune 2026 — Google Sheets collector
 *
 * Paste this over the existing Apps Script bound to your results spreadsheet.
 *
 * 1. Open the Google Sheet that already receives test results
 * 2. Extensions → Apps Script
 * 3. Delete the old code and paste this entire file
 * 4. Save
 * 5. Deploy → Manage deployments → the existing Web app → pencil →
 *    Version: New version → Deploy
 *    Keep the same URL so the website does not need to change
 *
 * After the next submission, the spreadsheet will contain:
 *   Responses          — one clean row per student
 *   Dashboard          — totals, charts, score bands, part averages
 *   Question Analysis  — % correct per question
 *   Demographics       — grade, school, transport, income, devices
 *
 * The answer key is used here only to score. It is not shown on the website.
 */

var KEYS = [
  "B","B","C","B","C",
  "B","C","D","B","C","B","C",
  "C","B","C","D","C","C","C","A",
  "C","A","B","C","D","B","B"
];

var MARKS = [
  0.5,0.5,0.5,0.5,0.5,
  1,1,1,1,1,1,1,
  1.5,1.5,1.5,1.5,1.5,1.5,1.5,1.5,
  0.5,0.5,0.5,0.5,0.5,0.5,0.5
];

var PARTS = [
  "A","A","A","A","A",
  "B","B","B","B","B","B","B",
  "C","C","C","C","C","C","C","C",
  "D","D","D","D","D","D","D"
];

var PART_MAX = { A: 2.5, B: 7, C: 12, D: 3.5 };

function doGet() {
  return ContentService
    .createTextOutput("FinLit Index collector is running.")
    .setMimeType(ContentService.MimeType.TEXT);
}

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    recordAttempt(data);
    return ContentService
      .createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function recordAttempt(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var scored = scorePayload(data);
  var responses = getOrCreateSheet_(ss, "Responses");
  ensureResponseHeaders_(responses);
  responses.appendRow(buildResponseRow_(data, scored));
  rebuildDashboard_(ss, responses);
}

function scorePayload(data) {
  var choices = [];
  var correctFlags = [];
  var parts = { A: 0, B: 0, C: 0, D: 0 };
  var total = 0;
  var correct = 0;
  var missed = [];

  for (var i = 0; i < 27; i++) {
    var choice = pickChoice_(data, i);
    var ok = choice && choice === KEYS[i];
    choices.push(choice || "blank");
    correctFlags.push(ok ? 1 : 0);
    if (ok) {
      parts[PARTS[i]] += MARKS[i];
      total += MARKS[i];
      correct++;
    } else {
      missed.push("Q" + (i + 1) + (choice ? " (" + choice + ")" : " (blank)"));
    }
  }

  return {
    choices: choices,
    correctFlags: correctFlags,
    parts: parts,
    total: Math.round(total * 100) / 100,
    correct: correct,
    missed: missed.length ? missed.join("; ") : "None — all correct"
  };
}

function pickChoice_(data, i) {
  var n = i + 1;
  var raw = data["q" + n + "_choice"] || data["q" + n] || "";
  raw = String(raw).trim();
  if (!raw || raw === "blank") return "";
  var letter = raw.match(/\b([ABCD])\b/);
  if (letter) return letter[1];
  if (raw === "OK") return KEYS[i];
  return "";
}

function responseHeaders_() {
  var headers = [
    "Timestamp", "Name", "Email", "Grade", "School",
    "Transport", "Parents occupation", "Area", "Devices", "Income",
    "Time used (min)", "Auto submitted", "Score / 25", "Correct / 27",
    "Part A / 2.5", "Part B / 7", "Part C / 12", "Part D / 3.5",
    "Missed questions"
  ];
  for (var i = 1; i <= 27; i++) {
    headers.push("Q" + i);
    headers.push("Q" + i + " correct");
  }
  return headers;
}

function ensureResponseHeaders_(sheet) {
  var headers = responseHeaders_();
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length)
      .setFontWeight("bold")
      .setBackground("#111111")
      .setFontColor("#ffffff")
      .setWrap(true);
    sheet.setFrozenRows(1);
    sheet.setFrozenColumns(5);
    sheet.setColumnWidth(1, 170);
    sheet.setColumnWidth(2, 160);
    sheet.setColumnWidth(5, 180);
    sheet.setColumnWidth(19, 280);
    return;
  }
  var existing = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  if (existing.join("|") !== headers.join("|")) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length)
      .setFontWeight("bold")
      .setBackground("#111111")
      .setFontColor("#ffffff");
  }
}

function buildResponseRow_(data, scored) {
  var timeMin = data.time_used_min;
  if (timeMin === undefined || timeMin === null || timeMin === "") {
    var parsed = String(data.time_used || "").replace(" min", "");
    timeMin = Number(parsed) || "";
  }
  var row = [
    data.submitted_at || new Date().toISOString(),
    data.name || "",
    data.email || "",
    data.grade || "",
    data.school || "",
    data.bg_transport || "",
    data.bg_parents || "",
    data.bg_area || "",
    data.bg_devices || "",
    data.bg_income || "",
    timeMin,
    data.auto_submitted || "",
    scored.total,
    scored.correct,
    scored.parts.A,
    scored.parts.B,
    scored.parts.C,
    scored.parts.D,
    scored.missed
  ];
  for (var i = 0; i < 27; i++) {
    row.push(scored.choices[i]);
    row.push(scored.correctFlags[i]);
  }
  return row;
}

function getOrCreateSheet_(ss, name) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  return sheet;
}

function rebuildDashboard_(ss, responses) {
  var data = readResponses_(responses);
  buildDashboard_(ss, data);
  buildQuestionAnalysis_(ss, data);
  buildDemographics_(ss, data);
}

function readResponses_(sheet) {
  var last = sheet.getLastRow();
  if (last < 2) {
    return { rows: [], n: 0 };
  }
  var values = sheet.getRange(2, 1, last - 1, sheet.getLastColumn()).getValues();
  var rows = values.filter(function (r) { return String(r[1]).trim() !== ""; });
  return { rows: rows, n: rows.length };
}

function buildDashboard_(ss, data) {
  var sheet = getOrCreateSheet_(ss, "Dashboard");
  sheet.clear();
  sheet.setHiddenGridlines(true);
  sheet.setColumnWidth(1, 220);
  sheet.setColumnWidth(2, 90);
  sheet.setColumnWidth(3, 90);
  sheet.setColumnWidth(4, 24);
  sheet.setColumnWidth(5, 180);
  sheet.setColumnWidth(6, 90);
  sheet.setColumnWidth(7, 90);

  sheet.getRange("A1").setValue("FinLit Index  ·  Pune 2026").setFontWeight("bold").setFontSize(18);
  sheet.getRange("A2").setValue("Updated automatically after every submission. Scores are out of 25.").setFontColor("#666666");

  var n = data.n;
  var scores = data.rows.map(function (r) { return Number(r[12]) || 0; });
  var correctCounts = data.rows.map(function (r) { return Number(r[13]) || 0; });
  var timedOut = data.rows.filter(function (r) {
    return String(r[11]).toLowerCase().indexOf("yes") !== -1;
  }).length;

  var kpis = [
    ["Attempts", n],
    ["Average score", n ? round1_(avg_(scores)) : 0],
    ["Median score", n ? round1_(median_(scores)) : 0],
    ["Highest score", n ? round1_(Math.max.apply(null, scores)) : 0],
    ["Lowest score", n ? round1_(Math.min.apply(null, scores)) : 0],
    ["Average correct", n ? round1_(avg_(correctCounts)) : 0],
    ["Timed out", timedOut]
  ];
  sheet.getRange("A4").setValue("Overview").setFontWeight("bold").setFontSize(13);
  sheet.getRange(5, 1, kpis.length, 2).setValues(kpis);
  sheet.getRange(5, 1, kpis.length, 1).setFontColor("#666666");
  sheet.getRange(5, 2, kpis.length, 1).setFontWeight("bold").setFontSize(14);

  var bands = [
    ["Score band", "Students"],
    ["0 to 5", countBand_(scores, 0, 5)],
    ["5+ to 10", countBand_(scores, 5, 10, true)],
    ["10+ to 15", countBand_(scores, 10, 15, true)],
    ["15+ to 20", countBand_(scores, 15, 20, true)],
    ["20 to 25", countBand_(scores, 20, 25)]
  ];
  sheet.getRange("A14").setValue("Score distribution").setFontWeight("bold").setFontSize(13);
  sheet.getRange(15, 1, bands.length, 2).setValues(bands);
  sheet.getRange(15, 1, 1, 2).setFontWeight("bold").setBackground("#111111").setFontColor("#ffffff");

  var gradeMap = groupAvg_(data.rows, 3, 12);
  var gradeTable = [["Grade", "Students", "Average score"]];
  ["9", "10", "11", "12"].forEach(function (g) {
    var item = gradeMap[g] || { count: 0, avg: 0 };
    gradeTable.push([g, item.count, item.count ? round1_(item.avg) : 0]);
  });
  sheet.getRange("A23").setValue("By grade").setFontWeight("bold").setFontSize(13);
  sheet.getRange(24, 1, gradeTable.length, 3).setValues(gradeTable);
  sheet.getRange(24, 1, 1, 3).setFontWeight("bold").setBackground("#111111").setFontColor("#ffffff");

  var partAvgs = [
    ["Part", "Average", "Maximum"],
    ["A Warm up", n ? round1_(avg_(data.rows.map(function (r) { return Number(r[14]) || 0; }))) : 0, PART_MAX.A],
    ["B Everyday", n ? round1_(avg_(data.rows.map(function (r) { return Number(r[15]) || 0; }))) : 0, PART_MAX.B],
    ["C Read carefully", n ? round1_(avg_(data.rows.map(function (r) { return Number(r[16]) || 0; }))) : 0, PART_MAX.C],
    ["D Market awareness", n ? round1_(avg_(data.rows.map(function (r) { return Number(r[17]) || 0; }))) : 0, PART_MAX.D]
  ];
  sheet.getRange("A31").setValue("Part averages").setFontWeight("bold").setFontSize(13);
  sheet.getRange(32, 1, partAvgs.length, 3).setValues(partAvgs);
  sheet.getRange(32, 1, 1, 3).setFontWeight("bold").setBackground("#111111").setFontColor("#ffffff");

  sheet.getCharts().forEach(function (c) { sheet.removeChart(c); });
  if (n === 0) return;

  addChart_(sheet, Charts.ChartType.COLUMN, sheet.getRange("A15:B20"), 4, 5, "Students in each score band");
  addChart_(sheet, Charts.ChartType.COLUMN, sheet.getRange("A24:C28"), 19, 5, "Average score by grade");
  addChart_(sheet, Charts.ChartType.BAR, sheet.getRange("A32:B36"), 32, 5, "Average marks by part");
}

function buildQuestionAnalysis_(ss, data) {
  var sheet = getOrCreateSheet_(ss, "Question Analysis");
  sheet.clear();
  sheet.setFrozenRows(1);
  var n = data.n;
  var headers = ["Question", "Part", "Marks", "Correct", "Attempts", "% correct", "Most common answer"];
  var rows = [headers];
  var chartRows = [["Question", "% correct"]];

  for (var i = 0; i < 27; i++) {
    var choiceCol = 19 + i * 2;
    var okCol = 20 + i * 2;
    var correct = 0;
    var counts = { A: 0, B: 0, C: 0, D: 0, blank: 0 };
    data.rows.forEach(function (r) {
      if (Number(r[okCol]) === 1) correct++;
      var ch = String(r[choiceCol] || "blank");
      if (counts[ch] === undefined) counts[ch] = 0;
      counts[ch]++;
    });
    var top = "—";
    var topN = -1;
    Object.keys(counts).forEach(function (k) {
      if (counts[k] > topN) { topN = counts[k]; top = k; }
    });
    var pct = n ? Math.round((correct / n) * 1000) / 10 : 0;
    rows.push(["Q" + (i + 1), PARTS[i], MARKS[i], correct, n, pct, n ? top : "—"]);
    chartRows.push(["Q" + (i + 1), pct]);
  }

  sheet.getRange(1, 1, rows.length, headers.length).setValues(rows);
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight("bold")
    .setBackground("#111111")
    .setFontColor("#ffffff");
  sheet.getRange(1, 20, chartRows.length, 2).setValues(chartRows);
  sheet.getRange(1, 20, 1, 2).setFontWeight("bold").setBackground("#111111").setFontColor("#ffffff");
  sheet.hideColumns(20, 2);
  sheet.setColumnWidth(1, 90);
  sheet.setColumnWidth(7, 170);

  sheet.getCharts().forEach(function (c) { sheet.removeChart(c); });
  if (n === 0) return;
  addChart_(sheet, Charts.ChartType.COLUMN, sheet.getRange(1, 20, chartRows.length, 2), 1, 9, "% of students correct on each question");
}

function buildDemographics_(ss, data) {
  var sheet = getOrCreateSheet_(ss, "Demographics");
  sheet.clear();
  sheet.getRange("A1").setValue("Background patterns").setFontWeight("bold").setFontSize(16);
  sheet.getRange("A2").setValue("These tables refresh after every submission.").setFontColor("#666666");

  writeGroupTable_(sheet, 4, 1, "By school", groupAvg_(data.rows, 4, 12), ["School", "Students", "Average score"]);
  writeGroupTable_(sheet, 4, 5, "By transport", groupAvg_(data.rows, 5, 12), ["Transport", "Students", "Average score"]);
  writeGroupTable_(sheet, 4, 9, "By income", groupAvg_(data.rows, 9, 12), ["Income", "Students", "Average score"]);
  writeGroupTable_(sheet, 20, 1, "By area", groupAvg_(data.rows, 7, 12), ["Area", "Students", "Average score"]);

  var deviceCounts = {};
  data.rows.forEach(function (r) {
    String(r[8] || "").split(",").map(function (s) { return s.trim(); }).filter(Boolean).forEach(function (d) {
      deviceCounts[d] = (deviceCounts[d] || 0) + 1;
    });
  });
  var deviceTable = [["Device", "Students"]];
  Object.keys(deviceCounts).sort().forEach(function (k) {
    deviceTable.push([k, deviceCounts[k]]);
  });
  if (deviceTable.length === 1) deviceTable.push(["—", 0]);
  sheet.getRange(20, 5).setValue("Devices owned").setFontWeight("bold").setFontSize(12);
  sheet.getRange(21, 5, deviceTable.length, 2).setValues(deviceTable);
  sheet.getRange(21, 5, 1, 2).setFontWeight("bold").setBackground("#111111").setFontColor("#ffffff");

  sheet.getCharts().forEach(function (c) { sheet.removeChart(c); });
  if (data.n === 0) return;

  var schoolCount = Math.max(Object.keys(groupAvg_(data.rows, 4, 12)).length, 1);
  addChart_(sheet, Charts.ChartType.BAR, sheet.getRange(5, 1, 1 + schoolCount, 3), 4, 12, "Average score by school");
  var transportCount = Math.max(Object.keys(groupAvg_(data.rows, 5, 12)).length, 1);
  addChart_(sheet, Charts.ChartType.COLUMN, sheet.getRange(5, 5, 1 + transportCount, 2), 20, 12, "How students come to school");
}

function writeGroupTable_(sheet, row, col, title, grouped, headers) {
  sheet.getRange(row, col).setValue(title).setFontWeight("bold").setFontSize(12);
  var keys = Object.keys(grouped).sort();
  var table = [headers];
  if (!keys.length) table.push(["—", 0, 0]);
  keys.forEach(function (k) {
    table.push([k || "—", grouped[k].count, round1_(grouped[k].avg)]);
  });
  sheet.getRange(row + 1, col, table.length, 3).setValues(table);
  sheet.getRange(row + 1, col, 1, 3).setFontWeight("bold").setBackground("#111111").setFontColor("#ffffff");
}

function addChart_(sheet, type, range, row, col, title) {
  var chart = sheet.newChart()
    .setChartType(type)
    .addRange(range)
    .setOption("title", title)
    .setOption("legend", { position: "none" })
    .setOption("colors", ["#111111"])
    .setOption("backgroundColor", "#ffffff")
    .setOption("hAxis", { textStyle: { color: "#111111" }, titleTextStyle: { color: "#111111" } })
    .setOption("vAxis", { textStyle: { color: "#111111" }, minValue: 0 })
    .setOption("titleTextStyle", { color: "#111111", fontSize: 12, bold: true })
    .setPosition(row, col, 0, 0)
    .build();
  sheet.insertChart(chart);
}

function groupAvg_(rows, keyIndex, valueIndex) {
  var map = {};
  rows.forEach(function (r) {
    var key = String(r[keyIndex] || "").trim() || "Not given";
    if (!map[key]) map[key] = { count: 0, sum: 0 };
    map[key].count++;
    map[key].sum += Number(r[valueIndex]) || 0;
  });
  Object.keys(map).forEach(function (k) {
    map[k].avg = map[k].count ? map[k].sum / map[k].count : 0;
  });
  return map;
}

function avg_(arr) {
  if (!arr.length) return 0;
  return arr.reduce(function (a, b) { return a + b; }, 0) / arr.length;
}

function median_(arr) {
  if (!arr.length) return 0;
  var a = arr.slice().sort(function (x, y) { return x - y; });
  var mid = Math.floor(a.length / 2);
  return a.length % 2 ? a[mid] : (a[mid - 1] + a[mid]) / 2;
}

function round1_(n) {
  return Math.round(n * 10) / 10;
}

function countBand_(scores, from, to, exclusiveStart) {
  return scores.filter(function (s) {
    if (exclusiveStart) return s > from && s <= to;
    return s >= from && s <= to;
  }).length;
}
