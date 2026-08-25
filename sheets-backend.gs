/**
 * FinLit Index · Pune 2026 — Google Sheets collector
 *
 * Spreadsheet:
 * https://docs.google.com/spreadsheets/d/1cntlU1VebqoA2AZE9Vqa1VQ6st-jW8xYNZwVD-Yp044/edit
 *
 * INSTALL (once)
 * 1. Open that spreadsheet
 * 2. Extensions → Apps Script
 * 3. Replace ALL code with this file and Save
 * 4. Select setup from the function dropdown → Run
 *    Grant permissions when asked
 * 5. Deploy → Manage deployments → pencil on the Web app
 *    Version: New version → Deploy
 *
 * After that, every new test submission rebuilds the dashboard and charts.
 * You can also use FinLit → Rebuild dashboard from the spreadsheet menu.
 *
 * Q22 key is B. Profession and correct-count are not stored.
 */

var KEYS = [
  "B","B","C","B","C",
  "B","C","D","B","C","B","C",
  "C","B","C","D","C","C","C","A",
  "C","B","B","C","D","B","B"
];

var MARKS = [
  0.5,0.5,0.5,0.5,0.5,
  1,1,1,1,1,1,1,
  1.5,1.5,1.5,1.5,1.5,1.5,1.5,1.5,
  0.5,0.5,0.5,0.5,0.5,0.5,0.5
];

var PARTS = "AAAAABBBBBBBCCCCCCCCDDDDDDD".split("");
var PART_MAX = { A: 2.5, B: 7, C: 12, D: 3.5 };
var NAVY = "#0f2438";
var CREAM = "#f7f5f0";
var GOLD = "#a89468";
var WHITE = "#fffcf7";

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("FinLit")
    .addItem("Rebuild dashboard", "setup")
    .addToUi();
}

function doGet() {
  setup();
  return ContentService
    .createTextOutput("FinLit dashboard rebuilt.")
    .setMimeType(ContentService.MimeType.TEXT);
}

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    ingestAttempt_(data);
    rebuildAll_();
    return ContentService
      .createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  migrateOldData_(ss);
  rebuildAll_();
}

function rebuildAll_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var responses = getOrCreateSheet_(ss, "Responses");
  rescoreResponses_(responses);
  var data = readResponses_(responses);
  buildDashboard_(ss, data);
  buildQuestionAnalysis_(ss, data);
  buildDemographics_(ss, data);
  hideUnusedSheets_(ss);
}

function ingestAttempt_(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = getOrCreateSheet_(ss, "Responses");
  ensureResponseHeaders_(sheet);
  sheet.appendRow(buildResponseRow_(data, scoreChoices_(extractChoices_(data))));
}

function extractChoices_(data) {
  var choices = [];
  for (var i = 0; i < 27; i++) {
    choices.push(parseChoice_(data["q" + (i + 1) + "_choice"] || data["q" + (i + 1)], i));
  }
  return choices;
}

function parseChoice_(raw, index) {
  raw = String(raw || "").trim();
  if (!raw || raw === "blank") return "";
  if (raw === "OK") {
    // Older rows stored OK instead of the letter. Q22 used to be keyed as A.
    return index === 21 ? "A" : KEYS[index];
  }
  var letter = raw.match(/\b([ABCD])\b/);
  return letter ? letter[1] : "";
}

function scoreChoices_(choices) {
  var parts = { A: 0, B: 0, C: 0, D: 0 };
  var total = 0;
  var missed = [];
  for (var i = 0; i < 27; i++) {
    var choice = choices[i] || "";
    if (choice && choice === KEYS[i]) {
      parts[PARTS[i]] += MARKS[i];
      total += MARKS[i];
    } else {
      missed.push("Q" + (i + 1) + (choice ? " (" + choice + ")" : " (blank)"));
    }
  }
  return {
    choices: choices.map(function (c) { return c || "blank"; }),
    parts: parts,
    total: Math.round(total * 100) / 100,
    missed: missed.length ? missed.join("; ") : "None"
  };
}

function responseHeaders_() {
  var headers = [
    "Timestamp", "Name", "Email", "Grade", "School",
    "Transport", "Parents occupation", "Area", "Devices", "Income",
    "Time used (min)", "Timed out", "Score / 25",
    "Part A / 2.5", "Part B / 7", "Part C / 12", "Part D / 3.5",
    "Missed questions"
  ];
  for (var i = 1; i <= 27; i++) headers.push("Q" + i);
  return headers;
}

function ensureResponseHeaders_(sheet) {
  var headers = responseHeaders_();
  var width = Math.max(sheet.getLastColumn(), headers.length);
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  } else {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight("bold")
    .setBackground(NAVY)
    .setFontColor(WHITE)
    .setWrap(true);
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(5);
  sheet.setColumnWidth(1, 170);
  sheet.setColumnWidth(2, 160);
  sheet.setColumnWidth(5, 180);
  sheet.setColumnWidth(18, 280);
  if (width > headers.length) {
    sheet.deleteColumns(headers.length + 1, width - headers.length);
  }
}

function buildResponseRow_(data, scored) {
  var timeMin = data.time_used_min;
  if (timeMin === undefined || timeMin === null || timeMin === "") {
    timeMin = Number(String(data.time_used || "").replace(" min", "")) || "";
  }
  var timedOut = String(data.auto_submitted || "").toLowerCase().indexOf("yes") !== -1 ? "Yes" : "No";
  var row = [
    data.submitted_at || new Date().toISOString(),
    data.name || "",
    data.email || "",
    data.grade || "",
    data.school || "",
    data.bg_transport || data.transport || "",
    data.bg_parents || "",
    data.bg_area || "",
    data.bg_devices || "",
    data.bg_income || "",
    timeMin,
    timedOut,
    scored.total,
    scored.parts.A,
    scored.parts.B,
    scored.parts.C,
    scored.parts.D,
    scored.missed
  ];
  return row.concat(scored.choices);
}

function migrateOldData_(ss) {
  var names = ss.getSheets().map(function (s) { return s.getName(); });
  var source = null;
  if (names.indexOf("Responses") !== -1 && ss.getSheetByName("Responses").getLastRow() > 1) {
    source = ss.getSheetByName("Responses");
  } else if (names.indexOf("Sheet1") !== -1 && ss.getSheetByName("Sheet1").getLastRow() > 1) {
    source = ss.getSheetByName("Sheet1");
  }
  if (!source) {
    ensureResponseHeaders_(getOrCreateSheet_(ss, "Responses"));
    return;
  }

  var values = source.getDataRange().getValues();
  if (values.length < 2) {
    ensureResponseHeaders_(getOrCreateSheet_(ss, "Responses"));
    return;
  }

  var headers = values[0].map(function (h) { return String(h); });
  var looksNew = headers.indexOf("Score / 25") !== -1 && headers.indexOf("profession") === -1 && headers.indexOf("correct_count") === -1;
  var rows = [];
  for (var r = 1; r < values.length; r++) {
    var rec = {};
    headers.forEach(function (h, c) { rec[h] = values[r][c]; });
    if (!(rec.name || rec.Name || rec.q1 || rec.Q1)) continue;
    if (looksNew) {
      var choices = [];
      for (var i = 0; i < 27; i++) choices.push(parseChoice_(rec["Q" + (i + 1)], i));
      rows.push(buildResponseRow_({
        submitted_at: rec["Timestamp"] || rec.submitted_at,
        name: rec.Name || rec.name,
        email: rec.Email || rec.email || "",
        grade: rec.Grade || rec.grade,
        school: rec.School || rec.school,
        bg_transport: rec.Transport || rec.bg_transport,
        bg_parents: rec["Parents occupation"] || rec.bg_parents,
        bg_area: rec.Area || rec.bg_area,
        bg_devices: rec.Devices || rec.bg_devices,
        bg_income: rec.Income || rec.bg_income,
        time_used_min: rec["Time used (min)"] || rec.time_used,
        auto_submitted: rec["Timed out"] === "Yes" ? "yes" : rec.auto_submitted
      }, scoreChoices_(choices)));
    } else {
      var oldChoices = [];
      for (var j = 0; j < 27; j++) oldChoices.push(parseChoice_(rec["q" + (j + 1)], j));
      rows.push(buildResponseRow_({
        submitted_at: rec.submitted_at,
        name: rec.name,
        email: rec.email || "",
        grade: rec.grade,
        school: rec.school,
        bg_transport: rec.bg_transport,
        bg_parents: rec.bg_parents,
        bg_area: rec.bg_area,
        bg_devices: rec.bg_devices,
        bg_income: rec.bg_income,
        time_used: rec.time_used,
        auto_submitted: rec.auto_submitted
      }, scoreChoices_(oldChoices)));
    }
  }

  var dest = getOrCreateSheet_(ss, "Responses");
  dest.clear();
  ensureResponseHeaders_(dest);
  if (rows.length) dest.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
}

function rescoreResponses_(sheet) {
  ensureResponseHeaders_(sheet);
  var last = sheet.getLastRow();
  if (last < 2) return;
  var values = sheet.getRange(2, 1, last - 1, sheet.getLastColumn()).getValues();
  var headers = responseHeaders_();
  var out = values.map(function (row) {
    var choices = [];
    for (var i = 0; i < 27; i++) choices.push(parseChoice_(row[18 + i], i));
    var scored = scoreChoices_(choices);
    var data = {
      submitted_at: row[0],
      name: row[1],
      email: row[2],
      grade: row[3],
      school: row[4],
      bg_transport: row[5],
      bg_parents: row[6],
      bg_area: row[7],
      bg_devices: row[8],
      bg_income: row[9],
      time_used_min: row[10],
      auto_submitted: String(row[11]).toLowerCase().indexOf("yes") !== -1 ? "yes" : "no"
    };
    var built = buildResponseRow_(data, scored);
    while (built.length < headers.length) built.push("");
    return built.slice(0, headers.length);
  });
  sheet.getRange(2, 1, out.length, headers.length).setValues(out);
}

function readResponses_(sheet) {
  var last = sheet.getLastRow();
  if (last < 2) return { rows: [], n: 0 };
  var values = sheet.getRange(2, 1, last - 1, sheet.getLastColumn()).getValues();
  var rows = values.filter(function (r) { return String(r[1]).trim() !== ""; });
  return { rows: rows, n: rows.length };
}

function styleTitle_(range, size) {
  range.setFontFamily("Georgia")
    .setFontWeight("bold")
    .setFontColor(NAVY)
    .setFontSize(size);
}

function buildDashboard_(ss, data) {
  var sheet = getOrCreateSheet_(ss, "Dashboard");
  sheet.clear();
  sheet.setHiddenGridlines(true);
  sheet.setTabColor(GOLD);
  sheet.getRange("A1:L45").setBackground(CREAM);
  [220, 100, 110, 24, 200, 90, 90, 24, 160, 90].forEach(function (w, i) {
    sheet.setColumnWidth(i + 1, w);
  });

  styleTitle_(sheet.getRange("A1"), 22);
  sheet.getRange("A1").setValue("The Skyward Project  ·  FinLit Index");
  sheet.getRange("A2").setValue("Pune 2026  ·  scores out of 25  ·  Q22 key is B").setFontColor(GOLD).setFontStyle("italic");

  var n = data.n;
  var scores = data.rows.map(function (r) { return Number(r[12]) || 0; });
  var timedOut = data.rows.filter(function (r) {
    return String(r[11]).toLowerCase().indexOf("yes") !== -1;
  }).length;

  var kpis = [
    ["Attempts", n],
    ["Average score", n ? round1_(avg_(scores)) : 0],
    ["Median score", n ? round1_(median_(scores)) : 0],
    ["Highest", n ? round1_(Math.max.apply(null, scores)) : 0],
    ["Lowest", n ? round1_(Math.min.apply(null, scores)) : 0],
    ["Timed out", timedOut]
  ];
  sheet.getRange("A4").setValue("Overview").setFontWeight("bold").setFontColor(NAVY).setFontSize(13);
  sheet.getRange(5, 1, kpis.length, 2).setValues(kpis);
  sheet.getRange(5, 1, kpis.length, 1).setFontColor("#5c574f");
  sheet.getRange(5, 2, kpis.length, 1).setFontWeight("bold").setFontSize(16).setFontColor(NAVY);

  var bands = [
    ["Score band", "Students"],
    ["0–5", countBand_(scores, 0, 5)],
    ["5–10", countBand_(scores, 5, 10, true)],
    ["10–15", countBand_(scores, 10, 15, true)],
    ["15–20", countBand_(scores, 15, 20, true)],
    ["20–25", countBand_(scores, 20, 25)]
  ];
  sheet.getRange("A13").setValue("Score distribution").setFontWeight("bold").setFontColor(NAVY).setFontSize(13);
  sheet.getRange(14, 1, bands.length, 2).setValues(bands);
  paintHeader_(sheet.getRange(14, 1, 1, 2));

  var gradeTable = [["Grade", "Students", "Average"]];
  var gradeMap = groupAvg_(data.rows, 3, 12);
  ["9", "10", "11", "12"].forEach(function (g) {
    var item = gradeMap[g] || { count: 0, avg: 0 };
    gradeTable.push([g, item.count, item.count ? round1_(item.avg) : 0]);
  });
  sheet.getRange("A22").setValue("By grade").setFontWeight("bold").setFontColor(NAVY).setFontSize(13);
  sheet.getRange(23, 1, gradeTable.length, 3).setValues(gradeTable);
  paintHeader_(sheet.getRange(23, 1, 1, 3));

  var partAvgs = [
    ["Part", "Average", "Maximum"],
    ["A  Warm up", n ? round1_(avg_(col_(data.rows, 13))) : 0, PART_MAX.A],
    ["B  Everyday", n ? round1_(avg_(col_(data.rows, 14))) : 0, PART_MAX.B],
    ["C  Read carefully", n ? round1_(avg_(col_(data.rows, 15))) : 0, PART_MAX.C],
    ["D  Market awareness", n ? round1_(avg_(col_(data.rows, 16))) : 0, PART_MAX.D]
  ];
  sheet.getRange("A30").setValue("Part averages").setFontWeight("bold").setFontColor(NAVY).setFontSize(13);
  sheet.getRange(31, 1, partAvgs.length, 3).setValues(partAvgs);
  paintHeader_(sheet.getRange(31, 1, 1, 3));

  sheet.getCharts().forEach(function (c) { sheet.removeChart(c); });
  addChart_(sheet, Charts.ChartType.COLUMN, sheet.getRange("A14:B19"), 4, 5, "Students in each score band");
  addChart_(sheet, Charts.ChartType.COLUMN, sheet.getRange("A23:C27"), 18, 5, "Average score by grade");
  addChart_(sheet, Charts.ChartType.BAR, sheet.getRange("A31:B35"), 31, 5, "Average marks by part");
}

function buildQuestionAnalysis_(ss, data) {
  var sheet = getOrCreateSheet_(ss, "Question Analysis");
  sheet.clear();
  sheet.setTabColor(NAVY);
  sheet.setFrozenRows(1);
  var n = data.n;
  var table = [["Question", "Part", "Marks", "Correct answer", "Got it right", "Attempts", "% correct", "Most common pick"]];
  var chartRows = [["Question", "% correct"]];

  for (var i = 0; i < 27; i++) {
    var right = 0;
    var counts = { A: 0, B: 0, C: 0, D: 0, blank: 0 };
    data.rows.forEach(function (r) {
      var ch = parseChoice_(r[18 + i], i) || "blank";
      counts[ch] = (counts[ch] || 0) + 1;
      if (ch === KEYS[i]) right++;
    });
    var top = "—", topN = -1;
    Object.keys(counts).forEach(function (k) {
      if (counts[k] > topN) { topN = counts[k]; top = k; }
    });
    var pct = n ? Math.round((right / n) * 1000) / 10 : 0;
    table.push(["Q" + (i + 1), PARTS[i], MARKS[i], KEYS[i], right, n, pct, n ? top : "—"]);
    chartRows.push(["Q" + (i + 1), pct]);
  }

  sheet.getRange(1, 1, table.length, 8).setValues(table);
  paintHeader_(sheet.getRange(1, 1, 1, 8));
  sheet.getRange(1, 10, chartRows.length, 2).setValues(chartRows);
  sheet.hideColumns(10, 2);
  sheet.setColumnWidth(1, 90);
  sheet.setColumnWidth(8, 160);
  sheet.getRange("A1:H1").setNote("Correct answers are for evaluators. They are not shown on the student website.");

  sheet.getCharts().forEach(function (c) { sheet.removeChart(c); });
  addChart_(sheet, Charts.ChartType.COLUMN, sheet.getRange(1, 10, chartRows.length, 2), 1, 9, "% of students correct on each question");
}

function buildDemographics_(ss, data) {
  var sheet = getOrCreateSheet_(ss, "Demographics");
  sheet.clear();
  sheet.setTabColor(GOLD);
  sheet.getRange("A1:P40").setBackground(CREAM);
  styleTitle_(sheet.getRange("A1"), 18);
  sheet.getRange("A1").setValue("Background patterns");
  sheet.getRange("A2").setValue("Refreshes after every submission and whenever you run Rebuild dashboard.").setFontColor("#5c574f");

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
  Object.keys(deviceCounts).sort().forEach(function (k) { deviceTable.push([k, deviceCounts[k]]); });
  if (deviceTable.length === 1) deviceTable.push(["—", 0]);
  sheet.getRange(20, 5).setValue("Devices owned").setFontWeight("bold").setFontColor(NAVY);
  sheet.getRange(21, 5, deviceTable.length, 2).setValues(deviceTable);
  paintHeader_(sheet.getRange(21, 5, 1, 2));

  sheet.getCharts().forEach(function (c) { sheet.removeChart(c); });
  var schoolN = Math.max(Object.keys(groupAvg_(data.rows, 4, 12)).length, 1);
  var transportN = Math.max(Object.keys(groupAvg_(data.rows, 5, 12)).length, 1);
  addChart_(sheet, Charts.ChartType.BAR, sheet.getRange(5, 1, 1 + schoolN, 3), 4, 12, "Average score by school");
  addChart_(sheet, Charts.ChartType.COLUMN, sheet.getRange(5, 5, 1 + transportN, 2), 20, 12, "How students come to school");
}

function writeGroupTable_(sheet, row, col, title, grouped, headers) {
  sheet.getRange(row, col).setValue(title).setFontWeight("bold").setFontColor(NAVY).setFontSize(12);
  var keys = Object.keys(grouped).sort();
  var table = [headers];
  if (!keys.length) table.push(["—", 0, 0]);
  keys.forEach(function (k) {
    table.push([k || "—", grouped[k].count, round1_(grouped[k].avg)]);
  });
  sheet.getRange(row + 1, col, table.length, 3).setValues(table);
  paintHeader_(sheet.getRange(row + 1, col, 1, 3));
}

function paintHeader_(range) {
  range.setFontWeight("bold").setBackground(NAVY).setFontColor(WHITE);
}

function addChart_(sheet, type, range, row, col, title) {
  var chart = sheet.newChart()
    .setChartType(type)
    .addRange(range)
    .setOption("title", title)
    .setOption("legend", { position: "none" })
    .setOption("colors", [NAVY])
    .setOption("backgroundColor", WHITE)
    .setOption("hAxis", { textStyle: { color: NAVY } })
    .setOption("vAxis", { textStyle: { color: NAVY }, minValue: 0 })
    .setOption("titleTextStyle", { color: NAVY, fontSize: 12, bold: true })
    .setPosition(row, col, 0, 0)
    .build();
  sheet.insertChart(chart);
}

function hideUnusedSheets_(ss) {
  var keep = { Responses: 1, Dashboard: 1, "Question Analysis": 1, Demographics: 1 };
  ss.getSheets().forEach(function (s) {
    if (!keep[s.getName()] && ss.getSheets().length > 1) {
      try { s.hideSheet(); } catch (e) {}
    }
  });
}

function getOrCreateSheet_(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
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

function col_(rows, i) {
  return rows.map(function (r) { return Number(r[i]) || 0; });
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
