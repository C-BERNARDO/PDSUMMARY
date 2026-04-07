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
  results.innerHTML = ""
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
    <div class="ant-upload-list-item">
      <span class="ant-upload-list-item-icon">📄</span>
      <span class="ant-upload-list-item-name" title="${escHtml(f.name)}">${escHtml(f.name)}</span>
      <button class="ant-upload-list-item-remove" onclick="removeFile(${i})" title="Remove">✕</button>
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
  return str.replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;")
}

function showStatus(message, type = "info") {
  const icons = { info: "⟳", success: "✓", error: "✕", warning: "⚠" }
  const spinClass = type === "info" ? " ant-loading-icon" : ""
  statusEl.className = `ant-alert ant-alert-${type === "processing" ? "info" : type}`
  statusEl.innerHTML = `
    <span class="ant-alert-icon${spinClass}">${icons[type === "processing" ? "info" : type] || "ℹ"}</span>
    <span class="ant-alert-message">${escHtml(message)}</span>
  `
  statusEl.classList.remove("hidden")

  if (window._stTimeout) clearTimeout(window._stTimeout)
  if (type !== "processing" && type !== "info") {
    window._stTimeout = setTimeout(() => hideStatus(), 4000)
  }
}

function hideStatus() {
  statusEl.classList.add("hidden")
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
  const h = Math.floor(s / 3600)
  const m = Math.floor((s % 3600) / 60)
  const sec = Math.floor(s % 60)
  return `${String(h).padStart(2,"0")}:${String(m).padStart(2,"0")}:${String(sec).padStart(2,"0")}`
}

// ── Core processing ──
async function processFiles() {
  showStatus("Processing files…", "processing")
  results.innerHTML = ""

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
        const wb = window.XLSX.read(new Uint8Array(e.target.result), { type: "array", cellDates: false, cellText: false })
        const ws = wb.Sheets[wb.SheetNames[0]]
        resolve(window.XLSX.utils.sheet_to_json(ws, { defval: "", raw: true, dateNF: "HH:MM:SS" }))
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
        const lk = k.toLowerCase().replace(/\s/g, "")
        for (const rkey of rk) {
          const lrkey = rkey.toLowerCase().replace(/\s/g, "")
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

  Object.values(consolidated).forEach(d => {
    const tc = d["Total Calls"]
    d["AVG Talk Time"]  = fmtTime(tc > 0 ? d["Talk Time"]  / tc : 0)
    d["AVG Wait Time"]  = fmtTime(tc > 0 ? d["Wait Time"]  / tc : 0)
    d["AVG Write Time"] = fmtTime(tc > 0 ? d["Write Time"] / tc : 0)
    d["Spent Time"]  = fmtTime(d["Spent Time"])
    d["Talk Time"]   = fmtTime(d["Talk Time"])
    d["Wait Time"]   = fmtTime(d["Wait Time"])
    d["Write Time"]  = fmtTime(d["Write Time"])
    d["Pause Time"]  = fmtTime(d["Pause Time"])
  })

  return consolidated
}

// ── Render results ──
function displayResults(consolidated) {
  const names = Object.keys(consolidated).sort()

  if (!names.length) {
    results.innerHTML = `
      <div class="ant-card">
        <div class="ant-card-body">
          <div class="ant-empty">
            <div class="ant-empty-image">📭</div>
            <div class="ant-empty-description">No collector data found</div>
          </div>
        </div>
      </div>`
    return
  }

  const cols = [
    { key: "Total Calls",    label: "Total Calls",  num: true, type: "int"  },
    { key: "Spent Time",     label: "Spent Time",   num: true, type: "time" },
    { key: "Talk Time",      label: "Talk Time",    num: true, type: "time" },
    { key: "AVG Talk Time",  label: "Avg Talk",     num: true, type: "avg"  },
    { key: "Wait Time",      label: "Wait Time",    num: true, type: "time" },
    { key: "AVG Wait Time",  label: "Avg Wait",     num: true, type: "avg"  },
    { key: "Write Time",     label: "Write Time",   num: true, type: "time" },
    { key: "AVG Write Time", label: "Avg Write",    num: true, type: "avg"  },
    { key: "Pause Time",     label: "Pause Time",   num: true, type: "time" },
    { key: "Pause Count",    label: "Pause Count",  num: true, type: "int"  },
  ]

  // ── Build summary row values ──
  const totals = {}
  let grandTotalCalls = 0

  cols.forEach(col => {
    if (col.type === "int") {
      totals[col.key] = names.reduce((sum, n) => sum + (consolidated[n][col.key] || 0), 0)
      if (col.key === "Total Calls") grandTotalCalls = totals[col.key]
    } else if (col.type === "time") {
      // Parse formatted strings back to seconds and sum
      const secs = names.reduce((sum, n) => sum + parseTime(consolidated[n][col.key] || "00:00:00"), 0)
      totals[col.key] = fmtTime(secs)
    } else if (col.type === "avg") {
      // Weighted average: sum of (avg * calls) / total calls
      const totalSecs = names.reduce((sum, n) => {
        const tc = consolidated[n]["Total Calls"] || 0
        return sum + parseTime(consolidated[n][col.key] || "00:00:00") * tc
      }, 0)
      totals[col.key] = grandTotalCalls > 0 ? fmtTime(totalSecs / grandTotalCalls) : "00:00:00"
    }
  })

  results.innerHTML = `
    <div class="ant-card">
      <div class="ant-card-head">
        <span class="ant-card-head-title">Consolidated Report</span>
        <span class="ant-card-extra">
          <span class="ant-tag ant-tag-blue">${names.length} Collectors</span>
          &nbsp;
          <span class="ant-tag ant-tag-green">${grandTotalCalls.toLocaleString()} Total Calls</span>
        </span>
      </div>
      <div class="ant-table-wrapper">
        <div class="ant-table-wrapper-inner">
          <table class="ant-table">
            <thead>
              <tr>
                <th>Collector</th>
                ${cols.map(c => `<th class="${c.num ? "num" : ""}">${c.label}</th>`).join("")}
              </tr>
            </thead>
            <tbody>
              ${names.map(name => {
                const d = consolidated[name]
                return `
                  <tr>
                    <td>${escHtml(name)}</td>
                    ${cols.map(c => `<td class="${c.num ? "num" : ""}">${d[c.key] ?? ""}</td>`).join("")}
                  </tr>`
              }).join("")}
            </tbody>
            <tfoot>
              <tr class="summary-row">
                <td class="summary-label">Grand Total</td>
                ${cols.map(c => `<td class="num summary-cell">${totals[c.key] ?? ""}</td>`).join("")}
              </tr>
            </tfoot>
          </table>
        </div>
        <div class="ant-table-footer">
          <span>Showing ${names.length} record${names.length !== 1 ? "s" : ""}</span>
          <span>Generated on ${processedTime}</span>
        </div>
      </div>
    </div>
  `
}