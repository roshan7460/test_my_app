let columns = [];
let currentRows = [];

// 🔑 Originals for RESET
let originalColumns = [];
let originalRows = [];

/* =========================
   Upload Excel
========================= */
function uploadExcel() {
    const file = document.getElementById("fileInput").files[0];
    if (!file) {
        alert("Select an Excel file first");
        return;
    }

    const formData = new FormData();
    formData.append("file", file);

    fetch("/upload", {
        method: "POST",
        body: formData
    })
    .then(res => res.json())
    .then(data => {
        columns = data.headers.map((h, i) => ({
            name: h,
            index: i,
            selected: true
        }));

        // 🔑 Save originals (deep copy)
        originalColumns = JSON.parse(JSON.stringify(columns));
        originalRows = JSON.parse(JSON.stringify(data.rows));

        currentRows = JSON.parse(JSON.stringify(data.rows));

        loadPreferences();
        renderTable();
    });
}

/* =========================
   Render Table
========================= */
function renderTable() {
    let html = "<table><tr>";

    // Header
    columns.forEach((col, i) => {
        html += `
            <th>
                <input type="checkbox" ${col.selected ? "checked" : ""}
                       onchange="toggleColumn(${i})"><br>

                <input type="text" value="${col.name}"
                       onchange="renameColumn(${i}, this.value)"
                       style="width:120px"><br>

                <button onclick="moveColumnUp(${i})">⬆</button>
                <button onclick="moveColumnDown(${i})">⬇</button>
            </th>
        `;
    });

    html += "<th>Action</th></tr>";

    // Rows
    currentRows.forEach((row, rIdx) => {
        const sumClass = row._isSumRow ? "sum-row" : "";
        html += `<tr class="fade-row ${sumClass}" id="row-${rIdx}">`;

        columns.forEach(col => {
            if (col.selected) {
                html += `<td>${row[col.index] ?? ""}</td>`;
            }
        });

        html += `
            <td>
                ${row._isSumRow ? "" : `<button class="delete-btn" onclick="removeRow(${rIdx})">❌</button>`}
            </td>
        </tr>`;
    });

    html += "</table>";
    document.getElementById("preview").innerHTML = html;

    updateSumDropdown();
    savePreferences();
}

/* =========================
   Column Controls
========================= */
function toggleColumn(i) {
    columns[i].selected = !columns[i].selected;
    renderTable();
}

function renameColumn(i, value) {
    columns[i].name = value;
    savePreferences();
}

function moveColumnUp(i) {
    if (i === 0) return;
    [columns[i - 1], columns[i]] = [columns[i], columns[i - 1]];
    renderTable();
}

function moveColumnDown(i) {
    if (i === columns.length - 1) return;
    [columns[i + 1], columns[i]] = [columns[i], columns[i + 1]];
    renderTable();
}

/* =========================
   Select / Deselect All
========================= */
function selectAllColumns() {
    columns.forEach(c => c.selected = true);
    renderTable();
}

function deselectAllColumns() {
    columns.forEach(c => c.selected = false);
    renderTable();
}

/* =========================
   Remove Row
========================= */
function removeRow(index) {
    const rowEl = document.getElementById(`row-${index}`);
    rowEl.classList.add("fade-out");

    setTimeout(() => {
        currentRows.splice(index, 1);
        renderTable();
    }, 300);
}

/* =========================
   SUM FEATURE (SAME ROW)
========================= */
function updateSumDropdown() {
    const select = document.getElementById("sumColumnSelect");
    if (!select) return;

    select.innerHTML = "";

    columns.forEach(col => {
        if (col.selected) {
            const opt = document.createElement("option");
            opt.value = col.index;
            opt.textContent = col.name;
            select.appendChild(opt);
        }
    });

    if (select.options.length === 0) {
        const opt = document.createElement("option");
        opt.textContent = "No column selected";
        select.appendChild(opt);
    }
}

function calculateSum() {
    const select = document.getElementById("sumColumnSelect");
    if (!select || !select.value) {
        alert("Please select a column");
        return;
    }

    const excelColIndex = Number(select.value);
    let sum = 0;
    let hasNumber = false;

    currentRows.forEach(row => {
        if (row._isSumRow) return;

        const val = row[excelColIndex];
        if (val !== "" && val !== null && !isNaN(val)) {
            sum += Number(val);   // calculation only
            hasNumber = true;
        }
    });

    if (!hasNumber) {
        alert("No numeric values found in this column");
        return;
    }

    let sumRow = currentRows.find(row => row._isSumRow);

    if (!sumRow) {
        sumRow = Array(columns.length).fill("");
        sumRow._isSumRow = true;
        currentRows.push(sumRow);
    }

    sumRow[excelColIndex] = sum;
    renderTable();
}

/* =========================
   RESET ALL (CLEAR HISTORY)
========================= */
function resetAll() {
    if (!confirm("This will reset all changes. Continue?")) return;

    columns = JSON.parse(JSON.stringify(originalColumns));
    currentRows = JSON.parse(JSON.stringify(originalRows));

    localStorage.removeItem("excel_pdf_columns");

    renderTable();
}

/* =========================
   Generate PDF
========================= */
function generatePDF() {
    const selectedCols = columns.filter(c => c.selected);

    if (selectedCols.length === 0) {
        alert("Please select at least one column");
        return;
    }

    const headers = selectedCols.map(c => c.name);
    const rows = currentRows.map(row =>
        selectedCols.map(c => row[c.index])
    );

    const btn = document.querySelector(".pdf-btn");
    btn.classList.add("loading");

    fetch("/generate_pdf", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ headers, rows })
    })
    .then(res => res.json())
    .then(data => {
        btn.classList.remove("loading");
        window.location.href = `/download/${data.pdf}`;
    });
}

/* =========================
   Preferences (localStorage)
========================= */
function savePreferences() {
    localStorage.setItem("excel_pdf_columns", JSON.stringify(columns));
}

function loadPreferences() {
    const saved = localStorage.getItem("excel_pdf_columns");
    if (saved) {
        try {
            const parsed = JSON.parse(saved);
            if (parsed.length === columns.length) {
                columns = parsed;
            }
        } catch {}
    }
}
