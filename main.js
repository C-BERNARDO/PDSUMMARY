let selectedFiles = []
let processedTime = ""

// ── DOM refs ──
const fileInput       = document.getElementById("fileInput")
const processBtn      = document.getElementById("processBtn")
const clearBtn        = document.getElementById("clearBtn")
const fileListItems   = document.getElementById("fileListItems")
const fileListWrapper = document.getElementById("fileListWrapper")
const fileCountBadge  = document.getElementById("fileCount")
const statusEl        = document.getElementById("status")
const results         = document.getElementById("results")
const topbarMeta      = document.getElementById("topbar-meta")
const dropArea        = document.getElementById("dropArea")

// ── Drag & drop ──
;["dragenter", "dragover", "dragleave", "drop"].forEach(evt =>
  dropArea.addEventListener(evt, e => { e.preventDefault(); e.stopPropagation() }, false)
)
;["dragenter", "dragover"].forEach(evt =>
  dropArea.addEventListener(evt, () => dropArea.classList.add("drag-over"), false)
)
;["dragleave", "drop"].forEach(evt =>
  dropArea.addEventListener(evt, () => dropArea.classList.remove("drag-over"), false)
)

dropArea.addEventListener("drop", e => {
  const files = e.dataTransfer.files
  fileInput.files = files
  selectedFiles = Array.from(files)
  displayFileList()
  syncProcessBtn()
}, false)

fileInput.addEventListener("change", e => {
  selectedFiles = Array.from(e.target.files)
  displayFileList()
  syncProcessBtn()
})

clearBtn.addEventListener("click", () => {
  selectedFiles = []
  fileInput.value = ""
  fileListItems.innerHTML = ""
  fileListWrapper.style.display = "none"
  hideStatus()
  topbarMeta.innerHTML = ""
  results.innerHTML = `
    <div class="empty-state">
      <div class="empty-icon">
        <svg viewBox="0 0 80 80" fill="none">
          <circle cx="40" cy="40" r="38" stroke="#e2e8f0" stroke-width="2"/>
          <rect x="24" y="28" width="32" height="4" rx="2" fill="#cbd5e0"/>
          <rect x="24" y="38" width="24" height="4" rx="2" fill="#e2e8f0"/>
          <rect x="24" y="48" width="28" height="4" rx="2" fill="#e2e8f0"/>
        </svg>
      </div>
      <h3 class="empty-title">No report generated yet</h3>
      <p class="empty-desc">Upload one or more Excel campaign summary files from the left panel, then click <strong>Process Files</strong> to generate the consolidated performance report.</p>
    </div>`
  syncProcessBtn()
})

processBtn.addEventListener("click", processFiles)

// ── Helpers ──
function syncProcessBtn() {
  processBtn.disabled = selectedFiles.length === 0
}

function displayFileList() {
  if (!selectedFiles.length) {
    fileListWrapper.style.display = "none"
    return
  }
  fileListWrapper.style.display = "block"
  fileCountBadge.textContent = selectedFiles.length
  fileListItems.innerHTML = selectedFiles.map((f, i) => `
    <div class="file-item">
      <svg class="file-item-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5">
        <path d="M9 12h6m-6 4h6m2 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z"/>
      </svg>
      <span class="file-item-name" title="${escHtml(f.name)}">${escHtml(f.name)}</span>
      <button class="file-item-remove" onclick="removeFile(${i})" title="Remove">✕</button>
    </div>
  `).join("")
}

function removeFile(index) {
  selectedFiles.splice(index, 1)
  displayFileList()
  syncProcessBtn()
  const dt = new DataTransfer()
  selectedFiles.forEach(f => dt.items.add(f))
  fileInput.files = dt.files
}

function escHtml(str) {
  return String(str)
    .replace(/&/g,"&amp;").replace(/</g,"&lt;")
    .replace(/>/g,"&gt;").replace(/"/g,"&quot;")
}

function showStatus(message, type = "info") {
  const map = { info: ["⟳", "spin"], success: ["✓",""], error: ["✕",""], warning: ["⚠",""] }
  const t = type === "processing" ? "info" : type
  const [icon, cls] = map[t] || ["ℹ",""]
  statusEl.className = `status-alert ${t}`
  statusEl.innerHTML = `
    <span class="status-icon ${cls}">${icon}</span>
    <span>${escHtml(message)}</span>
  `
  if (window._stT) clearTimeout(window._stT)
  if (t !== "info") window._stT = setTimeout(() => hideStatus(), 4500)
}

function hideStatus() {
  statusEl.className = "status-alert hidden"
}

// ── Time helpers ──
function parseTime(t) {
  if (!t || t === "00:00:00" || t === "") return 0
  if (typeof t === "number") return Math.round(t * 24 * 3600)
  const parts = String(t).trim().split(":").map(Number)
  if (parts.length === 3) return parts[0] * 3600 + parts[1] * 60 + parts[2]
  if (parts.length === 2) return parts[0] * 60 + parts[1]
  return 0
}

function fmtTime(s) {
  s = Math.max(0, Math.round(s))
  const h = Math.floor(s / 3600)
  const m = Math.floor((s % 3600) / 60)
  const sec = s % 60
  return `${String(h).padStart(2,"0")}:${String(m).padStart(2,"0")}:${String(sec).padStart(2,"0")}`
}

function fmtPct(n) {
  return isFinite(n) ? n.toFixed(2) + "%" : "—"
}

// ── Core processing ──
async function processFiles() {
  showStatus("Processing files…", "processing")
  results.innerHTML = ""
  topbarMeta.innerHTML = ""

  try {
    const allData = []
    for (const file of selectedFiles) {
      showStatus(`Reading ${file.name}…`, "processing")
      const rows = await readExcelFile(file)
      allData.push(...rows)
    }
    if (!allData.length) { showStatus("No data found in the uploaded files.", "error"); return }

    const consolidated = consolidateData(allData)
    if (!Object.keys(consolidated).length) { showStatus("No valid collector data found.", "error"); return }

    processedTime = new Date().toLocaleString()
    displayResults(consolidated)
    showStatus(`Successfully processed ${selectedFiles.length} file(s)`, "success")
  } catch (err) {
    showStatus(`Error: ${err.message}`, "error")
    console.error(err)
  }
}

function readExcelFile(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader()
    reader.onload = e => {
      try {
        const wb = window.XLSX.read(new Uint8Array(e.target.result), { type:"array", cellDates:false, cellText:false })
        const ws = wb.Sheets[wb.SheetNames[0]]
        resolve(window.XLSX.utils.sheet_to_json(ws, { defval:"", raw:true, dateNF:"HH:MM:SS" }))
      } catch (err) { reject(err) }
    }
    reader.onerror = () => reject(new Error(`Failed to read: ${file.name}`))
    reader.readAsArrayBuffer(file)
  })
}

function consolidateData(allData) {
  const consolidated = {}
  const nameKeys = ["Collector Name","CollectorName","Collector","collector name","Name","name","Agent","Agent Name","AgentName"]

  allData.forEach(row => {
    let name = ""
    for (const k of nameKeys) {
      if (row[k] && String(row[k]).trim()) { name = String(row[k]).trim(); break }
    }
    if (!name) return

    if (!consolidated[name]) {
      consolidated[name] = {
        "Total Calls": 0, "Spent Time": 0, "Talk Time": 0,
        "Wait Time": 0, "Write Time": 0, "Pause Time": 0, "Pause Count": 0,
      }
    }

    const get = (...keys) => {
      for (const k of keys) {
        if (row[k] !== undefined && row[k] !== null && row[k] !== "") return row[k]
      }
      const rk = Object.keys(row)
      for (const k of keys) {
        const lk = k.toLowerCase().replace(/\s/g,"")
        for (const rkey of rk) {
          const lrkey = rkey.toLowerCase().replace(/\s/g,"")
          if (lrkey.includes(lk) || lk.includes(lrkey)) {
            if (row[rkey] !== undefined && row[rkey] !== null && row[rkey] !== "") return row[rkey]
          }
        }
      }
      return null
    }

    consolidated[name]["Total Calls"]  += parseInt(get("Total Calls","TotalCalls","Calls","total calls")) || 0
    consolidated[name]["Spent Time"]   += parseTime(get("Spent Time","SpentTime","spent time","Spent") || "00:00:00")
    consolidated[name]["Talk Time"]    += parseTime(get("Talk Time","TalkTime","talk time","Talk") || "00:00:00")
    consolidated[name]["Wait Time"]    += parseTime(get("Wait Time","WaitTime","wait time","Wait") || "00:00:00")
    consolidated[name]["Write Time"]   += parseTime(get("Write Time","WriteTime","write time","Write","ACW") || "00:00:00")
    consolidated[name]["Pause Time"]   += parseTime(get("Pause Time","PauseTime","pause time","Pause") || "00:00:00")
    consolidated[name]["Pause Count"]  += parseInt(get("Pause Count","PauseCount","pause count","Pauses")) || 0
  })

  // Format — keep raw seconds for avg/occ calculations before formatting
  Object.values(consolidated).forEach(d => {
    const tc = d["Total Calls"]
    // Averages (raw seconds)
    d["_avg_talk"]  = tc > 0 ? d["Talk Time"]  / tc : 0
    d["_avg_wait"]  = tc > 0 ? d["Wait Time"]  / tc : 0
    d["_avg_write"] = tc > 0 ? d["Write Time"] / tc : 0

    // Occupancy Rate = (Talk Time + Write Time) / Spent Time
    d["_occ"] = d["Spent Time"] > 0
      ? ((d["Talk Time"] + d["Write Time"]) / d["Spent Time"]) * 100
      : null

    // Format times
    d["AVG Talk Time"]  = fmtTime(d["_avg_talk"])
    d["AVG Wait Time"]  = fmtTime(d["_avg_wait"])
    d["AVG Write Time"] = fmtTime(d["_avg_write"])
    d["Spent Time"]  = fmtTime(d["Spent Time"])
    d["Talk Time"]   = fmtTime(d["Talk Time"])
    d["Wait Time"]   = fmtTime(d["Wait Time"])
    d["Write Time"]  = fmtTime(d["Write Time"])
    d["Pause Time"]  = fmtTime(d["Pause Time"])
    d["Occ Rate"]    = d["_occ"] !== null ? fmtPct(d["_occ"]) : "—"
  })

  return consolidated
}

// ── Render results ──
function displayResults(consolidated) {
  const names = Object.keys(consolidated).sort()
  if (!names.length) {
    results.innerHTML = `<div class="empty-state"><div class="empty-title">No data found</div></div>`
    return
  }

  const cols = [
    { key: "Total Calls",    label: "Total Calls",   type: "int"  },
    { key: "Spent Time",     label: "Spent Time",    type: "time" },
    { key: "Talk Time",      label: "Talk Time",     type: "time" },
    { key: "AVG Talk Time",  label: "Avg Talk",      type: "avg", srcKey: "Talk Time"  },
    { key: "Wait Time",      label: "Wait Time",     type: "time" },
    { key: "AVG Wait Time",  label: "Avg Wait",      type: "avg", srcKey: "Wait Time"  },
    { key: "Write Time",     label: "Write Time",    type: "time" },
    { key: "AVG Write Time", label: "Avg Write",     type: "avg", srcKey: "Write Time" },
    { key: "Pause Time",     label: "Pause Time",    type: "time" },
    { key: "Pause Count",    label: "Pause Count",   type: "int"  },
  ]

  // ── Grand Total ──
  const totals = {}
  let grandCalls = 0

  cols.forEach(col => {
    if (col.type === "int") {
      totals[col.key] = names.reduce((s, n) => s + (consolidated[n][col.key] || 0), 0)
      if (col.key === "Total Calls") grandCalls = totals[col.key]
    } else if (col.type === "time") {
      const secs = names.reduce((s, n) => s + parseTime(consolidated[n][col.key] || "00:00:00"), 0)
      totals[col.key] = fmtTime(secs)
      totals[col.key + "_secs"] = secs   // keep raw for occ
    } else if (col.type === "avg") {
      const totalSecs = names.reduce((s, n) => {
        return s + parseTime(consolidated[n][col.key] || "00:00:00") * (consolidated[n]["Total Calls"] || 0)
      }, 0)
      totals[col.key] = grandCalls > 0 ? fmtTime(totalSecs / grandCalls) : "00:00:00"
    }
  })

  // ── Grand Occupancy Rate = (Talk Time + Write Time) / Spent Time ──
  const grandSpentSecs = totals["Spent Time_secs"] || 0
  const grandTalkSecs  = totals["Talk Time_secs"]  || 0
  const grandWriteSecs = totals["Write Time_secs"] || 0
  const grandOcc = grandSpentSecs > 0
    ? ((grandTalkSecs + grandWriteSecs) / grandSpentSecs) * 100
    : null
  const grandOccStr = grandOcc !== null ? fmtPct(grandOcc) : "—"

  // ── Topbar meta ──
  topbarMeta.innerHTML = `
    <span class="tag tag-blue">${names.length} Collector${names.length !== 1 ? "s" : ""}</span>
    <span class="tag tag-green">${grandCalls.toLocaleString()} Calls</span>
    <span class="tag tag-purple">OCC ${grandOccStr}</span>
    <span class="tag tag-amber">${processedTime}</span>
  `

  // ── Build table HTML ──
  const thCols = cols.map(c => `<th class="r">${c.label}</th>`).join("")
  const totalCells = cols.map(c => `<th class="r">${totals[c.key] ?? ""}</th>`).join("")

  const rows = names.map(name => {
    const d = consolidated[name]
    return `<tr>
      <td>${escHtml(name)}</td>
      ${cols.map(c => `<td class="r">${d[c.key] ?? ""}</td>`).join("")}
      <td class="r">${d["Occ Rate"] ?? "—"}</td>
    </tr>`
  }).join("")

  results.innerHTML = `
    <div class="report-card">
      <div class="report-card-head">
        <span class="report-card-title">Performance Report</span>
        <div class="report-card-tags">
          <span class="tag tag-blue">${names.length} Collector${names.length !== 1 ? "s" : ""}</span>
          <span class="tag tag-green">${grandCalls.toLocaleString()} Total Calls</span>
          <span class="tag tag-purple">OCC ${grandOccStr}</span>
        </div>
      </div>

      <div class="table-scroll">
        <table class="data-table">
          <thead>
            <!-- Occupancy Rate row -->
            <tr class="occ-row">
              <th>Occupancy Rate</th>
              ${cols.map(c => {
                if (c.key === "Spent Time")  return `<th class="r">${totals["Spent Time"]}</th>`
                if (c.key === "Talk Time")   return `<th class="r">${totals["Talk Time"]}</th>`
                if (c.key === "Write Time")  return `<th class="r">${totals["Write Time"]}</th>`
                return `<th class="r">—</th>`
              }).join("")}
              <th class="r" style="font-size:14px;">${grandOccStr}</th>
            </tr>

            <!-- Grand Total row -->
            <tr class="total-row">
              <th>Grand Total</th>
              ${totalCells}
              <th class="r">—</th>
            </tr>

            <!-- Column header row -->
            <tr class="col-header">
              <th>Collector</th>
              ${thCols}
              <th class="r">Occ. Rate</th>
            </tr>
          </thead>
          <tbody>
            ${rows}
          </tbody>
        </table>
      </div>

      <div class="table-foot">
        <span>${names.length} record${names.length !== 1 ? "s" : ""} · Occupancy Rate = (Talk Time + Write Time) ÷ Spent Time</span>
        <span>Generated ${processedTime}</span>
      </div>
    </div>
  `
}