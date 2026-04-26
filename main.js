"use strict";

var selectedFiles = [];
var processedTime = "";

// DOM refs
var fileInput       = document.getElementById("fileInput");
var processBtn      = document.getElementById("processBtn");
var clearBtn        = document.getElementById("clearBtn");
var fileListItems   = document.getElementById("fileListItems");
var fileListWrapper = document.getElementById("fileListWrapper");
var fileCountBadge  = document.getElementById("fileCount");
var statusEl        = document.getElementById("status");
var results         = document.getElementById("results");
var dropArea        = document.getElementById("dropArea");

// Drag and drop
["dragenter","dragover","dragleave","drop"].forEach(function(evt) {
  dropArea.addEventListener(evt, function(e){ e.preventDefault(); e.stopPropagation(); }, false);
});
["dragenter","dragover"].forEach(function(evt) {
  dropArea.addEventListener(evt, function(){ dropArea.classList.add("drag-over"); }, false);
});
["dragleave","drop"].forEach(function(evt) {
  dropArea.addEventListener(evt, function(){ dropArea.classList.remove("drag-over"); }, false);
});

dropArea.addEventListener("drop", function(e) {
  var files = e.dataTransfer.files;
  selectedFiles = Array.from(files);
  displayFileList();
  syncProcessBtn();
}, false);

fileInput.addEventListener("change", function(e) {
  selectedFiles = Array.from(e.target.files);
  displayFileList();
  syncProcessBtn();
});

clearBtn.addEventListener("click", function() {
  selectedFiles = [];
  fileInput.value = "";
  fileListItems.innerHTML = "";
  fileListWrapper.style.display = "none";
  hideStatus();
  results.innerHTML = getEmptyStateHtml();
  syncProcessBtn();
});

processBtn.addEventListener("click", processFiles);

function syncProcessBtn() {
  processBtn.disabled = selectedFiles.length === 0;
}

function getEmptyStateHtml() {
  return '<div class="empty-state">' +
    '<div class="empty-icon">' +
      '<svg viewBox="0 0 80 80" fill="none">' +
        '<circle cx="40" cy="40" r="38" stroke="#e2e8f0" stroke-width="2"/>' +
        '<rect x="24" y="28" width="32" height="4" rx="2" fill="#cbd5e0"/>' +
        '<rect x="24" y="38" width="24" height="4" rx="2" fill="#e2e8f0"/>' +
        '<rect x="24" y="48" width="28" height="4" rx="2" fill="#e2e8f0"/>' +
      '</svg>' +
    '</div>' +
    '<h3 class="empty-title">No report generated yet</h3>' +
    '<p class="empty-desc">Upload one or more Excel campaign summary files from the left panel, then click <strong>Process Files</strong> to generate the consolidated performance report.</p>' +
  '</div>';
}

function displayFileList() {
  if (!selectedFiles.length) {
    fileListWrapper.style.display = "none";
    return;
  }
  fileListWrapper.style.display = "block";
  fileCountBadge.textContent = selectedFiles.length;
  var html = "";
  for (var i = 0; i < selectedFiles.length; i++) {
    html += '<div class="file-item">' +
      '<svg class="file-item-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5">' +
        '<path d="M9 12h6m-6 4h6m2 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z"/>' +
      '</svg>' +
      '<span class="file-item-name" title="' + escHtml(selectedFiles[i].name) + '">' + escHtml(selectedFiles[i].name) + '</span>' +
      '<button class="file-item-remove" onclick="removeFile(' + i + ')" title="Remove">x</button>' +
    '</div>';
  }
  fileListItems.innerHTML = html;
}

function removeFile(index) {
  selectedFiles.splice(index, 1);
  displayFileList();
  syncProcessBtn();
  var dt = new DataTransfer();
  for (var i = 0; i < selectedFiles.length; i++) dt.items.add(selectedFiles[i]);
  fileInput.files = dt.files;
}

function escHtml(str) {
  return String(str)
    .replace(/&/g,"&amp;")
    .replace(/</g,"&lt;")
    .replace(/>/g,"&gt;")
    .replace(/"/g,"&quot;");
}

function showStatus(message, type) {
  if (!type) type = "info";
  var t = type === "processing" ? "info" : type;
  statusEl.className = "status-alert " + t;
  statusEl.innerHTML = '<span>' + escHtml(message) + '</span>';
  if (window._stT) clearTimeout(window._stT);
}

function hideStatus() {
  statusEl.className = "status-alert hidden";
}

// Time helpers
function parseTime(t) {
  if (!t || t === "00:00:00" || t === "") return 0;
  if (typeof t === "number") return Math.round(t * 24 * 3600);
  var parts = String(t).trim().split(":").map(Number);
  if (parts.length === 3) return parts[0] * 3600 + parts[1] * 60 + parts[2];
  if (parts.length === 2) return parts[0] * 60 + parts[1];
  return 0;
}

function fmtTime(s) {
  s = Math.max(0, Math.round(s));
  var h = Math.floor(s / 3600);
  var m = Math.floor((s % 3600) / 60);
  var sec = s % 60;
  return pad2(h) + ":" + pad2(m) + ":" + pad2(sec);
}

function pad2(n) {
  return n < 10 ? "0" + n : "" + n;
}

function fmtPct(n) {
  return isFinite(n) ? n.toFixed(2) + "%" : "--";
}

// Core processing
function processFiles() {
  showStatus("Processing files...", "processing");
  results.innerHTML = "";

  if (!window.XLSX) {
    showStatus("ERROR: Excel library (XLSX) not loaded. Check your internet connection and reload the page.", "error");
    return;
  }

  var promise = Promise.resolve([]);
  for (var fi = 0; fi < selectedFiles.length; fi++) {
    promise = (function(file, prevPromise) {
      return prevPromise.then(function(allData) {
        showStatus("Reading " + file.name + "...", "processing");
        return readExcelFile(file).then(function(rows) {
          return allData.concat(rows);
        });
      });
    })(selectedFiles[fi], promise);
  }

  promise.then(function(allData) {
    if (!allData.length) {
      showStatus("No data found in the uploaded files.", "error");
      return;
    }
    var consolidated = consolidateData(allData);
    var keys = Object.keys(consolidated);
    if (!keys.length) {
      showStatus("No valid collector data found. Ensure your file has a Collector Name or Agent column.", "error");
      return;
    }
    processedTime = new Date().toLocaleString();
    displayResults(consolidated);
    showStatus("Done! Processed " + selectedFiles.length + " file(s), found " + keys.length + " agent(s).", "success");
  }).catch(function(err) {
    var msg = err && err.message ? err.message : String(err);
    showStatus("ERROR: " + msg, "error");
    console.error("processFiles error:", err);
  });
}

function readExcelFile(file) {
  return new Promise(function(resolve, reject) {
    var reader = new FileReader();
    reader.onload = function(e) {
      try {
        var data = new Uint8Array(e.target.result);
        var wb = XLSX.read(data, { type: "array", cellDates: false, cellText: false });
        var ws = wb.Sheets[wb.SheetNames[0]];
        var rows = XLSX.utils.sheet_to_json(ws, { defval: "", raw: true, dateNF: "HH:MM:SS" });
        resolve(rows);
      } catch(err) {
        reject(new Error("Failed to parse " + file.name + ": " + err.message));
      }
    };
    reader.onerror = function() {
      reject(new Error("Failed to read file: " + file.name));
    };
    reader.readAsArrayBuffer(file);
  });
}

function getVal(row, keys) {
  for (var i = 0; i < keys.length; i++) {
    var k = keys[i];
    if (row[k] !== undefined && row[k] !== null && row[k] !== "") return row[k];
  }
  var rowKeys = Object.keys(row);
  for (var i = 0; i < keys.length; i++) {
    var lk = keys[i].toLowerCase().replace(/\s/g, "");
    for (var j = 0; j < rowKeys.length; j++) {
      var lrk = rowKeys[j].toLowerCase().replace(/\s/g, "");
      if (lrk.includes(lk) || lk.includes(lrk)) {
        if (row[rowKeys[j]] !== undefined && row[rowKeys[j]] !== null && row[rowKeys[j]] !== "") {
          return row[rowKeys[j]];
        }
      }
    }
  }
  return null;
}

function consolidateData(allData) {
  var consolidated = {};
  var nameKeys = ["Collector Name","CollectorName","Collector","collector name","Name","name","Agent","Agent Name","AgentName"];

  for (var ri = 0; ri < allData.length; ri++) {
    var row = allData[ri];
    var name = "";
    for (var ki = 0; ki < nameKeys.length; ki++) {
      if (row[nameKeys[ki]] && String(row[nameKeys[ki]]).trim()) {
        name = String(row[nameKeys[ki]]).trim();
        break;
      }
    }
    if (!name) continue;

    if (!consolidated[name]) {
      consolidated[name] = {
        "Total Calls": 0,
        "Spent Time": 0,
        "Talk Time": 0,
        "Wait Time": 0,
        "Write Time": 0,
        "Pause Time": 0,
        "Pause Count": 0
      };
    }

    var d = consolidated[name];
    d["Total Calls"]  += parseInt(getVal(row, ["Total Calls","TotalCalls","Calls","total calls"])) || 0;
    d["Spent Time"]   += parseTime(getVal(row, ["Spent Time","SpentTime","spent time","Spent"]) || "00:00:00");
    d["Talk Time"]    += parseTime(getVal(row, ["Talk Time","TalkTime","talk time","Talk"]) || "00:00:00");
    d["Wait Time"]    += parseTime(getVal(row, ["Wait Time","WaitTime","wait time","Wait"]) || "00:00:00");
    d["Write Time"]   += parseTime(getVal(row, ["Write Time","WriteTime","write time","Write","ACW"]) || "00:00:00");
    d["Pause Time"]   += parseTime(getVal(row, ["Pause Time","PauseTime","pause time","Pause"]) || "00:00:00");
    d["Pause Count"]  += parseInt(getVal(row, ["Pause Count","PauseCount","pause count","Pauses"])) || 0;
  }

  var names = Object.keys(consolidated);
  for (var ni = 0; ni < names.length; ni++) {
    var d = consolidated[names[ni]];
    var tc = d["Total Calls"];

    var avgTalkSecs  = tc > 0 ? d["Talk Time"]  / tc : 0;
    var avgWaitSecs  = tc > 0 ? d["Wait Time"]  / tc : 0;
    var avgWriteSecs = tc > 0 ? d["Write Time"] / tc : 0;

    var occPct = d["Spent Time"] > 0
      ? ((d["Talk Time"] + d["Write Time"]) / d["Spent Time"]) * 100
      : null;

    d["AVG Talk Time"]  = fmtTime(avgTalkSecs);
    d["AVG Wait Time"]  = fmtTime(avgWaitSecs);
    d["AVG Write Time"] = fmtTime(avgWriteSecs);
    d["Occ Rate"]       = occPct !== null ? fmtPct(occPct) : "--";

    d["Spent Time"]  = fmtTime(d["Spent Time"]);
    d["Talk Time"]   = fmtTime(d["Talk Time"]);
    d["Wait Time"]   = fmtTime(d["Wait Time"]);
    d["Write Time"]  = fmtTime(d["Write Time"]);
    d["Pause Time"]  = fmtTime(d["Pause Time"]);
  }

  return consolidated;
}

function displayResults(consolidated) {
  var names = Object.keys(consolidated).sort();

  var cols = [
    { key: "Total Calls",    label: "Total Calls",  type: "int"  },
    { key: "Spent Time",     label: "Spent Time",   type: "time" },
    { key: "Talk Time",      label: "Talk Time",    type: "time" },
    { key: "AVG Talk Time",  label: "Avg Talk",     type: "avg"  },
    { key: "Wait Time",      label: "Wait Time",    type: "time" },
    { key: "AVG Wait Time",  label: "Avg Wait",     type: "avg"  },
    { key: "Write Time",     label: "Write Time",   type: "time" },
    { key: "AVG Write Time", label: "Avg Write",    type: "avg"  },
    { key: "Pause Time",     label: "Pause Time",   type: "time" },
    { key: "Pause Count",    label: "Pause Count",  type: "int"  }
  ];

  // Grand totals
  var totals = {};
  var grandCalls = 0;

  for (var ci = 0; ci < cols.length; ci++) {
    var col = cols[ci];
    if (col.type === "int") {
      var sum = 0;
      for (var ni = 0; ni < names.length; ni++) sum += consolidated[names[ni]][col.key] || 0;
      totals[col.key] = sum;
      if (col.key === "Total Calls") grandCalls = sum;
    } else if (col.type === "time") {
      var secs = 0;
      for (var ni = 0; ni < names.length; ni++) secs += parseTime(consolidated[names[ni]][col.key] || "00:00:00");
      totals[col.key] = fmtTime(secs);
      totals[col.key + "_s"] = secs;
    } else if (col.type === "avg") {
      var totalSecs = 0;
      for (var ni = 0; ni < names.length; ni++) {
        totalSecs += parseTime(consolidated[names[ni]][col.key] || "00:00:00") * (consolidated[names[ni]]["Total Calls"] || 0);
      }
      totals[col.key] = grandCalls > 0 ? fmtTime(totalSecs / grandCalls) : "00:00:00";
      totals[col.key + "_s"] = grandCalls > 0 ? totalSecs / grandCalls : 0;
    }
  }

  var grandSpentS = totals["Spent Time_s"] || 0;
  var grandTalkS  = totals["Talk Time_s"]  || 0;
  var grandWriteS = totals["Write Time_s"] || 0;
  var grandOcc    = grandSpentS > 0 ? ((grandTalkS + grandWriteS) / grandSpentS) * 100 : null;
  var grandOccStr = grandOcc !== null ? fmtPct(grandOcc) : "--";

  var totalAgents    = names.length;
  var avgSpentPerAgent = totalAgents > 0 ? grandSpentS / totalAgents : 0;
  var avgTalkPerAgent  = totalAgents > 0 ? grandTalkS  / totalAgents : 0;
  var avgWaitPerCall   = totals["AVG Wait Time_s"]  || 0;
  var avgWritePerCall  = totals["AVG Write Time_s"] || 0;

  // Summary table (single row of labels + single row of values)
  // Summary metrics
  var summaryLabels = [
    "Total Agent",
    "Total Calls",
    "Spent Time",
    "Avg. Spent Time",
    "Talk Time",
    "Talk Time / Agent",
    "Wait Time",
    "Wait Time / Call",
    "ACW",
    "ACW Per Call",
    "Pause Time",
    "Occupancy Rate"
  ];
  var summaryValues = [
    totalAgents.toLocaleString(),
    grandCalls.toLocaleString(),
    totals["Spent Time"],
    fmtTime(avgSpentPerAgent),
    totals["Talk Time"],
    fmtTime(avgTalkPerAgent),
    totals["Wait Time"],
    fmtTime(avgWaitPerCall),
    totals["Write Time"],
    fmtTime(avgWritePerCall),
    totals["Pause Time"],
    grandOccStr
  ];

  // Build the horizontal summary table header and value cells
  var sumHeaderCells = "";
  var sumValueCells  = "";
  for (var i = 0; i < summaryLabels.length; i++) {
    var isOcc  = i === summaryLabels.length - 1;
    var isPause = i === summaryLabels.length - 2;
    var cls = isOcc ? " sum-occ" : (isPause ? " sum-pause" : "");
    sumHeaderCells += '<th class="sum-th' + cls + '">' + summaryLabels[i] + '</th>';
    sumValueCells  += '<td class="sum-td' + cls + '">' + summaryValues[i]  + '</td>';
  }

  // Tab-separated value string for copy
  var copyText = summaryValues.join("\t");

  // Table header cells
  var thCols = "";
  for (var ci = 0; ci < cols.length; ci++) {
    thCols += '<th class="r">' + cols[ci].label + '</th>';
  }

  // Grand total row cells
  var totalCells = "";
  for (var ci = 0; ci < cols.length; ci++) {
    var v = totals[cols[ci].key];
    totalCells += '<th class="r">' + (v !== undefined && v !== null ? v : "") + '</th>';
  }

  // Data rows
  var rowsHtml = "";
  for (var ni = 0; ni < names.length; ni++) {
    var d = consolidated[names[ni]];
    var cells = "";
    for (var ci = 0; ci < cols.length; ci++) {
      var v = d[cols[ci].key];
      cells += '<td class="r">' + (v !== undefined && v !== null ? v : "") + '</td>';
    }
    rowsHtml +=
      '<tr>' +
        '<td>' + escHtml(names[ni]) + '</td>' +
        cells +
      '</tr>';
  }

  // Build full HTML
  var html =
    '<div class="report-wrapper">' +

      // Performance Summary horizontal table
      '<div class="summary-card">' +
        '<div class="summary-card-head">' +
          '<span class="summary-card-title">Performance Summary</span>' +
          '<button class="copy-btn" onclick="copySummary()">Copy Values</button>' +
        '</div>' +
        '<div class="summary-scroll">' +
          '<table class="summary-table">' +
            '<thead><tr>' + sumHeaderCells + '</tr></thead>' +
            '<tbody><tr id="summaryValueRow">' + sumValueCells + '</tr></tbody>' +
          '</table>' +
        '</div>' +
        '<div class="summary-copy-text" id="summaryText">' + copyText + '</div>' +
      '</div>' +

      // Collector results table (no header card, just the table)
      '<div class="results-table-wrap">' +
        '<div class="table-scroll">' +
          '<table class="data-table">' +
            '<thead>' +
              '<tr class="total-row"><th>Grand Total</th>' + totalCells + '</tr>' +
              '<tr class="col-header"><th>Collector</th>' + thCols + '</tr>' +
            '</thead>' +
            '<tbody>' + rowsHtml + '</tbody>' +
          '</table>' +
        '</div>' +
      '</div>' +

    '</div>';

  results.innerHTML = html;
}

function copySummary() {
  var el = document.getElementById("summaryText");
  if (!el) return;
  var text = el.textContent || el.innerText;
  var btn = document.querySelector(".copy-btn");
  function confirm() {
    if (btn) { btn.textContent = "Copied!"; btn.style.background = "#15803d"; }
    setTimeout(function() {
      if (btn) { btn.textContent = "Copy Values"; btn.style.background = ""; }
    }, 2000);
  }
  if (navigator.clipboard) {
    navigator.clipboard.writeText(text).then(confirm);
  } else {
    var ta = document.createElement("textarea");
    ta.value = text;
    ta.style.cssText = "position:fixed;opacity:0;top:0;left:0";
    document.body.appendChild(ta);
    ta.focus(); ta.select();
    document.execCommand("copy");
    document.body.removeChild(ta);
    confirm();
  }
}