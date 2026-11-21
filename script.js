const userAccounts = {
    "admin": { pass: "admin123", role: "admin", nvCode: "" },
    "HAILT60": { pass: "123456", role: "user", nvCode: "HAILT60" },
    "Thaonn11": { pass: "123456", role: "user", nvCode: "Thaonn11" }
};

// Tai khoan nhan su DVKH: ten dang nhap = ten nhan su, pass chung 19006600.
// dvkh khien nguoi dung chi xem duoc hop dong thuoc "Nhân sự DVKH" cua minh.
const DVKH_PASSWORD = "19006600";
[
    "ACCT.NGAVTN", "CANHPC2", "CHIPTH", "DUNGDTM7", "GIANGTN7", "KHANHNH43",
    "LAMLTT7", "LINHNPP", "MAINTN26", "NYDTY2", "PHUONGLND3", "PHUONGMTM2",
    "QUYNHNHN4", "THUYLTP8", "TIENPTM2", "TRANGCTT11", "TRUCNTT28", "TUDTB",
    "TUYENPTK4", "UYENPNL2", "XUANLTT7", "YENNTK16", "YENTTT9"
].forEach(name => {
    userAccounts[name] = { pass: DVKH_PASSWORD, role: "user", nvCode: "", dvkh: name };
});

const REPORT_COLUMNS = [
    "khu vực",
    "HD",
    "Tên KH",
    "Nhân sự DVKH",
    "Nhân viên thu",
    "Tình Trạng",
    "Kết quả đối soát",
    "Mức độ xử lý",
    "Số Ngày TT muộn so với T-1",
    "Tuổi thọ HD",
    "Trạng thái chăm sóc"
];

const GVIZ_URL = "https://docs.google.com/spreadsheets/d/1W2JbsnZr3fJTihZMro70HMhUWm1cYYSt5s26MVdwSHc/gviz/tq?tqx=out:json&sheet=BC";

let currentUser = null;
let globalData = [];
let refreshTimer = null;
let filterTimer = null;
let appliedFilters = { area: "", staff: "" };
let collectorOptions = [];
let selectedStaffs = [];

const IMPORT_STORAGE_KEY = "qlcn_assignment_import_v1";
let assignmentImport = null;   // {fileName, importedAt, map: {HD: {collector, dvkh}}}
try {
    const saved = localStorage.getItem(IMPORT_STORAGE_KEY);
    if (saved) {
        const parsed = JSON.parse(saved);
        if (parsed && parsed.map) assignmentImport = parsed;
    }
} catch (e) {
    assignmentImport = null;
}

function normImportText(value) {
    return String(value ?? "")
        .normalize("NFD").replace(/[\u0300-\u036f]/g, "")
        .replace(/đ/g, "d").replace(/Đ/g, "D")
        .toLowerCase().replace(/\s+/g, " ").trim();
}

function normImportHd(value) {
    if (typeof value === "number" && Number.isFinite(value)) {
        return String(value).replace(/\.0+$/, "").toUpperCase();
    }
    return String(value ?? "").trim().toUpperCase();
}

// Lay ngay trong thang tu ngay TT thang truoc (Date, serial Excel, "dd/mm", so 1-31)
function parsePrevPayDay(value) {
    if (value === null || value === undefined || value === "") return null;
    if (value instanceof Date && !isNaN(value)) return value.getDate();
    if (typeof value === "number" && Number.isFinite(value)) {
        if (value >= 1 && value <= 31) return Math.floor(value);
        if (value > 25569 && value < 80000) {
            const d = new Date(Math.round((value - 25569) * 86400 * 1000));
            return d.getUTCDate();
        }
        return null;
    }
    const text = String(value).trim();
    if (!text) return null;
    let m = text.match(/\d{4}[\/.-]\d{1,2}[\/.-](\d{1,2})/);   // yyyy-mm-dd -> ngay o cuoi
    if (m) return parseInt(m[1], 10);
    m = text.match(/(\d{1,2})\s*[\/.-]\s*\d{1,2}/);            // dd/mm hoac dd/mm/yyyy
    if (m) {
        const d = parseInt(m[1], 10);
        return d >= 1 && d <= 31 ? d : null;
    }
    m = text.match(/\b(\d{1,2})\b/);                          // so le "15"
    if (m) {
        const d = parseInt(m[1], 10);
        return d >= 1 && d <= 31 ? d : null;
    }
    return null;
}

// Cung nguong severity() trong auto_sync.py
function severityForLateDays(days) {
    if (days > 10) return "CỰC KỲ GẤP";
    if (days > 5) return "GẤP";
    if (days > 0) return "CẦN THEO DÕI";
    return "BÌNH THƯỜNG";
}

// Du lieu goc + anh xa phan cong tu file import (neu co)
function effectiveData() {
    if (!assignmentImport || !assignmentImport.map) return globalData;
    const map = assignmentImport.map;
    const todayDay = new Date().getDate();
    return globalData.map(item => {
        const m = map[normImportHd(item.HD)];
        if (!m) return item;
        const clone = { ...item };
        if (m.collector) clone["Nhân viên thu"] = m.collector;
        if (m.dvkh) clone["Nhân sự DVKH"] = m.dvkh;
        if (m.prevDay != null && isUnpaid(clone)) {
            // HD con no + da qua ngay TT thang truoc => tre todayDay - prevDay ngay
            const lateDays = Math.max(0, todayDay - m.prevDay);
            clone["Số Ngày TT muộn so với T-1"] = String(lateDays);
            clone["Mức độ xử lý"] = severityForLateDays(lateDays);
        }
        return clone;
    });
}

function findImportColumns(rows2d) {
    const HD_NAMES = ["so hd", "hop dong", "so hop dong", "ma hd", "contract", "hd"];
    const COL_NAMES = ["nhan vien thu", "nv thu", "phan cong", "nhan vien thu cuoc", "collector"];
    const DVKH_NAMES = ["nhan su dvkh", "dvkh", "nvvp au", "nhan su"];
    for (let r = 0; r < Math.min(rows2d.length, 10); r++) {
        const cells = rows2d[r].map(normImportText);
        const hd = cells.findIndex(c => HD_NAMES.some(w => c === w || c.includes("hop dong")));
        if (hd < 0) continue;
        const collector = cells.findIndex(c => COL_NAMES.some(w => c.includes(w)));
        const dvkh = cells.findIndex(c => DVKH_NAMES.some(w => c.includes(w)));
        // Cot ngay thanh toan thang T- (thang truoc)
        let prevDay = cells.findIndex(c =>
            (c.includes("ngay") || c.includes("tt")) &&
            (c.includes("t-") || c.includes("thang truoc") || c.includes("ky truoc")));
        if (prevDay < 0) {
            prevDay = cells.findIndex((c, i) =>
                i !== hd && i !== collector && i !== dvkh &&
                (c.includes("ngay thanh toan") || c.includes("ngay tt") || c.includes("ngay nop")));
        }
        return { headerRow: r, hd, collector, dvkh, prevDay };
    }
    return { headerRow: 0, hd: 0, collector: 1, dvkh: 2, prevDay: 3 };
}

function parseCsvText(text) {
    const clean = text.replace(/^\uFEFF/, "");
    const lines = clean.split(/\r?\n/).filter(line => line.trim() !== "");
    const delim = lines[0].includes("\t") ? "\t" : lines[0].includes(";") ? ";" : ",";
    return lines.map(line =>
        line.split(delim).map(cell => cell.trim().replace(/^"|"$/g, ""))
    );
}

async function parseImportRows(file) {
    const ext = file.name.split(".").pop().toLowerCase();
    if (ext === "csv") {
        return parseCsvText(await file.text());
    }
    if (typeof XLSX === "undefined") {
        throw new Error("Chưa tải được thư viện đọc Excel. Kiểm tra mạng rồi thử lại.");
    }
    const buf = await file.arrayBuffer();
    const wb = XLSX.read(buf, { type: "array", cellDates: true });
    const ws = wb.Sheets[wb.SheetNames[0]];
    return XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: "" });
}

function renderImportStatus() {
    const el = document.getElementById("import-status");
    if (!el) return;
    if (!assignmentImport) {
        el.hidden = true;
        el.innerHTML = "";
        return;
    }
    const count = Object.keys(assignmentImport.map).length;
    const d = new Date(assignmentImport.importedAt);
    const when = `${d.toLocaleDateString("vi-VN")} ${d.toLocaleTimeString("vi-VN", { hour: "2-digit", minute: "2-digit" })}`;
    el.hidden = false;
    el.innerHTML = `📄 ${escapeHtml(assignmentImport.fileName)} — ${count.toLocaleString("vi-VN")} HĐ — ${when} <a id="import-clear" title="Xóa dữ liệu đã import">Xóa</a>`;
    el.querySelector("#import-clear").addEventListener("click", clearAssignmentImport);
}

function clearAssignmentImport() {
    assignmentImport = null;
    localStorage.removeItem(IMPORT_STORAGE_KEY);
    renderImportStatus();
    if (globalData.length) {
        populateDropdowns(effectiveData());
        executeFilter();
    }
}

function downloadImportTemplate() {
    const headers = ["Số HĐ", "Nhân viên thu", "Nhân sự DVKH", "Ngày thanh toán T-"];
    const rows = [
        ["123456789", "NGUYEN VAN A", "TRAN THI B", 15],
        ["987654321", "NGUYEN VAN A", "LE VAN C", "20/10/2026"],
        ["555666777", "PHAM VAN D", "NGO THI E", 5],
    ];
    if (typeof XLSX !== "undefined") {
        const ws = XLSX.utils.aoa_to_sheet([headers, ...rows]);
        ws["!cols"] = [{ wch: 14 }, { wch: 18 }, { wch: 18 }, { wch: 18 }];
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, "Mau import");
        XLSX.writeFile(wb, "mau_import_thang_T.xlsx");
        return;
    }
    const csv = "\uFEFF" + [headers, ...rows].map(r => r.join(",")).join("\n");
    const a = document.createElement("a");
    a.href = URL.createObjectURL(new Blob([csv], { type: "text/csv;charset=utf-8" }));
    a.download = "mau_import_thang_T.csv";
    a.click();
    URL.revokeObjectURL(a.href);
}

async function handleImportFile(file) {
    const rows2d = await parseImportRows(file);
    const loc = findImportColumns(rows2d);
    const map = {};
    let withPrevDay = 0;
    for (let r = loc.headerRow + 1; r < rows2d.length; r++) {
        const row = rows2d[r];
        const hd = normImportHd(row[loc.hd]);
        if (!hd) continue;
        const prevDay = loc.prevDay >= 0 ? parsePrevPayDay(row[loc.prevDay]) : null;
        if (prevDay != null) withPrevDay++;
        map[hd] = {
            collector: loc.collector >= 0 ? String(row[loc.collector] || "").trim().toUpperCase() : "",
            dvkh: loc.dvkh >= 0 ? String(row[loc.dvkh] || "").trim().toUpperCase() : "",
            prevDay,
        };
    }
    const total = Object.keys(map).length;
    if (total === 0) {
        throw new Error("Không đọc được hợp đồng nào. File cần có cột: Số HĐ, Nhân viên thu, Nhân sự DVKH (và Ngày thanh toán T- nếu có).");
    }
    assignmentImport = { fileName: file.name, importedAt: new Date().toISOString(), map };
    localStorage.setItem(IMPORT_STORAGE_KEY, JSON.stringify(assignmentImport));
    renderImportStatus();
    if (globalData.length) {
        populateDropdowns(effectiveData());
        executeFilter();
    }
    const dateNote = withPrevDay
        ? `\n${withPrevDay.toLocaleString("vi-VN")} HĐ có ngày thanh toán T- → đã tính lại số ngày trễ.`
        : "\nKhông thấy cột Ngày thanh toán T- — giữ nguyên số ngày trễ từ hệ thống.";
    alert(`Đã import ${total.toLocaleString("vi-VN")} hợp đồng phân công tháng T từ file "${file.name}".` + dateNote);
}

function escapeHtml(value) {
    return String(value ?? "")
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;");
}

function gvizCell(cell) {
    if (!cell || cell.v == null || cell.v === "") return "";
    if (typeof cell.v === "string" && cell.v.indexOf("Date(") === 0) {
        return cell.f || cell.v;
    }
    return cell.v;
}

function isHeaderRow(values) {
    const normalized = values.map(normalizeHeaderName);
    return normalized.includes(normalizeHeaderName("khu vực"))
        && (normalized.includes(normalizeHeaderName("HD"))
            || normalized.includes(normalizeHeaderName("Hợp đồng")))
        && normalized.includes(normalizeHeaderName("Tên KH"));
}

function normalizeHeaderName(value) {
    return String(value ?? "").replace(/\s+/g, " ").trim().toLocaleLowerCase("vi");
}

function isUnpaid(item) {
    const reconciliation = String(item["Kết quả đối soát"] || "").trim().toLocaleLowerCase("vi");
    if (reconciliation) {
        return !/^(đã\s*(thanh toán|nộp tiền|thu|đóng)|đã\s*tt|paid|hoàn thành)$/.test(reconciliation);
    }

    const paymentStatus = String(item["Tình Trạng"] || "").trim().toLocaleLowerCase("vi");
    return !/đã\s*(thanh toán|nộp tiền|thu|đóng)|đã\s*tt|paid|hoàn thành/.test(paymentStatus);
}

function normalizeProcessingLevel(value) {
    const level = String(value ?? "").trim().toLowerCase();
    if (["cực gấp", "cực kỳ gấp", "cực gấp cần os xl ngay"].includes(level)) {
        return "CỰC KỲ GẤP";
    }
    if (level === "gấp") return "GẤP";
    if (level === "cần theo dõi") return "CẦN THEO DÕI";
    return level.toUpperCase();
}

function toReportRow(raw) {
    const item = {};
    REPORT_COLUMNS.forEach(key => {
        item[key] = raw[key] != null ? raw[key] : "";
    });
    if (!item["Mức độ xử lý"]) {
        const days = parseFloat(item["Số Ngày TT muộn so với T-1"]) || 0;
        if (days > 10) item["Mức độ xử lý"] = "CỰC KỲ GẤP";
        else if (days > 5) item["Mức độ xử lý"] = "GẤP";
        else item["Mức độ xử lý"] = "CẦN THEO DÕI";
    }
    item["Mức độ xử lý"] = normalizeProcessingLevel(item["Mức độ xử lý"]);
    return item;
}

async function fetchGoogleSheetRows() {
    const response = await fetch(GVIZ_URL + "&t=" + Date.now());
    if (!response.ok) {
        throw new Error("Không đọc được Google Sheet (" + response.status + ")");
    }
    const text = await response.text();
    const start = text.indexOf("{");
    const end = text.lastIndexOf("}");
    if (start === -1 || end === -1) {
        throw new Error("Google Sheet không trả về JSON. Hãy chia sẻ sheet ở chế độ xem bằng link.");
    }
    const payload = JSON.parse(text.slice(start, end + 1));
    if (payload.status !== "ok") {
        const message = payload.errors && payload.errors[0]
            ? payload.errors[0].detailed_message
            : "GViz trả về lỗi";
        throw new Error(message);
    }

    const cols = payload.table.cols || [];
    const rows = payload.table.rows || [];
    if (rows.length === 0) return [];

    const rowValues = row => cols.map((_, index) => gvizCell(row.c?.[index]));
    const firstRow = rowValues(rows[0]);
    const rowContainsHeaders = isHeaderRow(firstRow);
    const headers = rowContainsHeaders
        ? firstRow
        : cols.map(column => String(column.label || "").trim());
    const columnIndexMap = new Map();
    headers.forEach((header, index) => {
        const key = normalizeHeaderName(header);
        if (key) columnIndexMap.set(key, index);
    });

    const body = rowContainsHeaders ? rows.slice(1) : rows;
    return body.map(row => {
        const cells = row.c || [];
        const raw = {};
        REPORT_COLUMNS.forEach(columnName => {
            const columnIndex = columnIndexMap.get(normalizeHeaderName(columnName));
            raw[columnName] = columnIndex === undefined ? "" : gvizCell(cells[columnIndex]);
        });
        return toReportRow(raw);
    }).filter(item => REPORT_COLUMNS.some(key => String(item[key]).trim() !== ""));
}

function handleLogin() {
    const username = document.getElementById("username").value.trim();
    const password = document.getElementById("password").value.trim();
    const error = document.getElementById("login-error");
    const account = userAccounts[username];

    if (account && account.pass === password) {
        currentUser = { username, ...account };
        error.style.display = "none";
        document.getElementById("login-box").style.display = "none";
        document.getElementById("main-app").style.display = "block";
        document.getElementById("user-display").textContent = username;
        document.getElementById("role-display").textContent = currentUser.role === "admin"
            ? "(Quản trị viên)"
            : "(Nhân viên)";
        document.getElementById("filter-kv").value = "";
        document.getElementById("filter-dvkh").value = "";
        appliedFilters = { area: "", staff: "" };
        loadReport(true);
        return;
    }

    error.textContent = "Tên đăng nhập hoặc mật khẩu không chính xác!";
    error.style.display = "block";
}

function handleLogout() {
    currentUser = null;
    globalData = [];
    setActiveReportTab("summary");
    if (refreshTimer) {
        clearInterval(refreshTimer);
        refreshTimer = null;
    }
    document.getElementById("login-box").style.display = "block";
    document.getElementById("main-app").style.display = "none";
    document.getElementById("username").value = "";
    document.getElementById("password").value = "";
}

async function loadReport(applySelectedFilters = false) {
    const container = document.getElementById("table-container");
    const infoCount = document.getElementById("info-count");
    infoCount.textContent = "Đang nạp dữ liệu realtime từ Google Sheet...";

    try {
        globalData = await fetchGoogleSheetRows();
        populateDropdowns(effectiveData());
        if (applySelectedFilters) {
            appliedFilters = {
                area: document.getElementById("filter-kv").value,
                staff: document.getElementById("filter-dvkh").value
            };
        }
        executeFilter();
        if (!refreshTimer) {
            refreshTimer = setInterval(loadReport, 60000);
        }
    } catch (err) {
        infoCount.textContent = "";
        container.innerHTML = `<div style="color:red; padding: 12px;">Lỗi nạp dữ liệu: ${escapeHtml(err.message)}</div>`;
    }
}

function populateDropdowns(data) {
    fillSelect("filter-kv", uniqueValues(data, "khu vực"), "Menu thả xuống tìm khu vực");
    fillSelect("filter-dvkh", uniqueValues(data, "Nhân sự DVKH"), "Nhân sự DVKH");
    const statuses = [...new Set(data
        .map(item => String(item["Tình Trạng"] ?? "").trim())
        .filter(Boolean))]
        .sort((a, b) => a.localeCompare(b, "vi"));
    fillSelect("filter-status", statuses, "Tất cả tình trạng");

    // Khoa bo loc DVKH theo tai khoan dang nhap (neu co)
    const dvkhSelect = document.getElementById("filter-dvkh");
    if (currentUser && currentUser.dvkh) {
        const dvkhVal = String(currentUser.dvkh).trim().toUpperCase();
        if (![...dvkhSelect.options].some(option => option.value === dvkhVal)) {
            const option = document.createElement("option");
            option.value = dvkhVal;
            option.textContent = dvkhVal;
            dvkhSelect.appendChild(option);
        }
        dvkhSelect.value = dvkhVal;
        dvkhSelect.disabled = true;
        appliedFilters.staff = dvkhVal;
    } else {
        dvkhSelect.disabled = false;
    }

    collectorOptions = uniqueValues(data, "Nhân viên thu");
    if (currentUser && currentUser.role === "user" && currentUser.nvCode) {
        selectedStaffs = [String(currentUser.nvCode).trim().toUpperCase()];
    } else {
        selectedStaffs = selectedStaffs.filter(staff => collectorOptions.includes(staff));
    }

    renderCollectorOptions();

    if (window.jQuery && jQuery.fn.select2) {
        jQuery("#filter-kv, #filter-dvkh").trigger("change.select2");
    }
}

function renderCollectorOptions() {
    document.querySelectorAll("[data-collector-filter]").forEach(filter => {
        const container = filter.querySelector(".collector-filter-options");
        container.replaceChildren();

        collectorOptions.forEach(staff => {
            const label = document.createElement("label");
            label.className = "collector-filter-option";
            label.dataset.staffName = staff;

            const checkbox = document.createElement("input");
            checkbox.type = "checkbox";
            checkbox.value = staff;

            const name = document.createElement("span");
            name.textContent = staff;

            label.append(checkbox, name);
            container.appendChild(label);
        });

        filterCollectorOptions(filter);
    });

    syncCollectorControls();
}

function updateCollectorFilterLabel() {
    const selectedNames = selectedStaffs
        .filter(staff => collectorOptions.includes(staff))
        .sort((a, b) => a.localeCompare(b, "vi"));
    const label = selectedNames.length
        ? `${selectedNames.join(", ")} (${selectedNames.length} đã chọn)`
        : "Chọn nhân viên thu (0 đã chọn)";

    document.querySelectorAll("[data-collector-filter] .collector-filter-button").forEach(toggle => {
        toggle.textContent = label;
        toggle.title = label;
    });
}

function syncCollectorControls() {
    const isLocked = Boolean(currentUser && currentUser.role === "user" && currentUser.nvCode);
    document.querySelectorAll("[data-collector-filter]").forEach(filter => {
        filter.querySelectorAll(".collector-filter-options input[type='checkbox']").forEach(checkbox => {
            checkbox.checked = selectedStaffs.includes(checkbox.value);
            checkbox.disabled = isLocked;
        });

        const selectAll = filter.querySelector(".select-all-staff");
        selectAll.checked = collectorOptions.length > 0
            && collectorOptions.every(staff => selectedStaffs.includes(staff));
        selectAll.disabled = isLocked;
        filter.querySelector(".collector-filter-button").disabled = isLocked;
    });

    updateCollectorFilterLabel();
}

function filterCollectorOptions(filter) {
    const query = filter.querySelector(".collector-filter-search").value
        .trim()
        .toLowerCase();
    filter.querySelectorAll(".collector-filter-options .collector-filter-option").forEach(option => {
        const labelText = option.textContent.trim().toLowerCase();
        option.style.display = labelText.includes(query) ? "flex" : "none";
    });
}

function uniqueValues(data, key) {
    const values = data
        .map(item => String(item[key] ?? "").trim().toUpperCase())
        .filter(value => value !== "");
    return [...new Set(values)].sort((a, b) => a.localeCompare(b, "vi"));
}

function fillSelect(selectId, options, defaultText) {
    const select = document.getElementById(selectId);
    const previous = select.value;
    select.innerHTML = `<option value="">${escapeHtml(defaultText)}</option>`;
    options.forEach(option => {
        const element = document.createElement("option");
        element.value = option;
        element.textContent = option;
        select.appendChild(element);
    });
    if (previous && [...select.options].some(option => option.value === previous)) {
        select.value = previous;
    }
}

function scheduleFilter() {
    clearTimeout(filterTimer);
    filterTimer = setTimeout(executeFilter, 200);
}

function executeFilter() {
    if (globalData.length === 0) {
        renderSummary([]);
        renderTable([]);
        return;
    }

    const contractQuery = document.getElementById("filter-hd").value.toLowerCase().trim();
    const filterVal = (document.getElementById("selectMucDoXuLy")?.value || "").trim().toUpperCase();
    const statusFilter = (document.getElementById("filter-status")?.value || "").trim().toLocaleLowerCase("vi");
    const { area, staff } = appliedFilters;
    const allCollectorsSelected = collectorOptions.length > 0
        && collectorOptions.every(collector => selectedStaffs.includes(collector));

    const filtered = effectiveData().filter(item => {
        const contract = String(item.HD || "").toLowerCase();
        const matchContract = !contractQuery || contract.includes(contractQuery);
        const matchArea = !area || String(item["khu vực"] || "").trim().toUpperCase() === area.trim().toUpperCase();
        const matchStaff = !staff || String(item["Nhân sự DVKH"] || "").trim().toUpperCase() === staff.trim().toUpperCase();
        const collector = String(item["Nhân viên thu"] || "").trim().toUpperCase();
        const matchCollector = selectedStaffs.length === 0
            || allCollectorsSelected
            || selectedStaffs.includes(collector);
        return matchContract && matchArea && matchStaff && matchCollector;
    });

    renderSummary(filtered);
    const detailRows = filtered.filter(item => {
        if (!isUnpaid(item)) return false;
        const status = String(item["Tình Trạng"] ?? "").trim().toLocaleLowerCase("vi");
        if (statusFilter && status !== statusFilter) return false;
        const rowVal = normalizeProcessingLevel(item["Mức độ xử lý"]).toUpperCase();
        const selectedLevel = normalizeProcessingLevel(filterVal).toUpperCase();
        const matchMucDo = !filterVal || rowVal === selectedLevel;
        return matchMucDo;
    });
    renderTable(detailRows);
}

function renderSummary(data) {
    const container = document.getElementById("summary-container");
    const groups = new Map();

    const getGroup = collector => {
        if (!groups.has(collector)) {
            groups.set(collector, { assigned: 0, paid: 0, cared: 0, notCared: 0, late: 0 });
        }
        return groups.get(collector);
    };

    const countCareLate = (item, summary) => {
        const careStatus = String(item["Trạng thái chăm sóc"] || "").trim();
        if (careStatus === "Đã chăm sóc") summary.cared += 1;
        if (careStatus === "Chưa chăm sóc") summary.notCared += 1;
        const daysValue = item["Số Ngày TT muộn so với T-1"];
        if (daysValue !== "" && daysValue != null && Number.isFinite(Number(daysValue)) && Number(daysValue) > 0) {
            summary.late += 1;
        }
    };

    const importMap = assignmentImport && assignmentImport.map;
    if (importMap) {
        // Danh sach phan cong thang moi lay tu file import:
        // HĐ co trong import ma khong con trong du lieu ton => "Đã thanh toán"
        const unpaidByHd = new Map();
        data.forEach(item => {
            if (isUnpaid(item)) unpaidByHd.set(normImportHd(item.HD), item);
        });
        for (const [hd, m] of Object.entries(importMap)) {
            const collector = String(m.collector || "").trim() || "Chưa phân công";
            const summary = getGroup(collector);
            summary.assigned += 1;
            const row = unpaidByHd.get(hd);
            if (row) {
                countCareLate(row, summary);
            } else {
                summary.paid += 1;
            }
        }
        // HĐ con ton trong du lieu nhung chua co trong danh sach import
        data.forEach(item => {
            if (!isUnpaid(item) || importMap[normImportHd(item.HD)]) return;
            const collector = String(item["Nhân viên thu"] || "").trim() || "Chưa phân công";
            const summary = getGroup(collector);
            summary.assigned += 1;
            countCareLate(item, summary);
        });
    } else {
        data.forEach(item => {
            const collector = String(item["Nhân viên thu"] || "").trim() || "Chưa phân công";
            const summary = getGroup(collector);
            summary.assigned += 1;

            const paymentStatus = String(item["Tình Trạng"] || "").trim().toLocaleLowerCase("vi");
            const reconciliation = String(item["Kết quả đối soát"] || "").trim();
            if (reconciliation
                ? reconciliation === "Đã thanh toán"
                : /đã\s*(thanh toán|nộp tiền|thu|đóng)|đã\s*tt|paid|hoàn thành/.test(paymentStatus)) {
                summary.paid += 1;
            }

            countCareLate(item, summary);
        });
    }

    const headers = [
        "NHÂN VIÊN THU",
        "TỔNG PHÂN CÔNG",
        "TỔNG TỒN",
        "ĐÃ THANH TOÁN",
        "TỈ LỆ THANH TOÁN",
        "CHƯA CHĂM SÓC",
        "ĐÃ CHĂM SÓC",
        "TỈ LỆ CHĂM SÓC",
        "SO SÁNH T-1",
        "THANH TOÁN MUỘN"
    ];

    let html = "<h3>Báo cáo tổng hợp</h3><table id=\"summary-table\"><thead><tr>";
    headers.forEach((header, index) => {
        const title = index === 8
            ? "Tỷ lệ hợp đồng không thanh toán muộn so với T-1"
            : index === 9
                ? "Số hợp đồng có số ngày thanh toán muộn lớn hơn 0"
                : "";
        html += `<th${title ? ` title=\"${title}\"` : ""}>${escapeHtml(header)}</th>`;
    });
    html += "</tr></thead><tbody>";

    const rows = [...groups.entries()].sort((a, b) => a[0].localeCompare(b[0], "vi"));
    if (rows.length === 0) {
        html += `<tr><td colspan=\"${headers.length}\">Không có dữ liệu phù hợp với bộ lọc.</td></tr>`;
    } else {
        rows.forEach(([collector, summary]) => {
            const outstanding = summary.assigned - summary.paid;
            const paidRate = summary.assigned ? summary.paid / summary.assigned * 100 : 0;
            const caredRate = summary.assigned ? summary.cared / summary.assigned * 100 : 0;
            const values = [
                collector,
                summary.assigned,
                outstanding,
                summary.paid,
                `${paidRate.toFixed(1)}%`,
                summary.notCared,
                summary.cared,
                `${caredRate.toFixed(1)}%`,
                `${(((summary.assigned - summary.late) / summary.assigned) * 100).toFixed(1)}%`,
                summary.late
            ];
            html += "<tr>" + values.map(value => `<td>${escapeHtml(value)}</td>`).join("") + "</tr>";
        });
    }

    html += "</tbody></table>";
    container.innerHTML = html;
}

function renderTable(data) {
    const container = document.getElementById("table-container");
    const infoCount = document.getElementById("info-count");
    infoCount.textContent = `Tổng số bản ghi hiển thị: ${data.length}`;

    if (data.length === 0) {
        container.innerHTML = "<p style='padding: 10px;'>Không tìm thấy dữ liệu phù hợp với bộ lọc!</p>";
        return;
    }

    const importMap = assignmentImport && assignmentImport.map;
    const showPrevDay = Boolean(importMap) && data.some(row => {
        const m = importMap[normImportHd(row.HD)];
        return m && m.prevDay != null;
    });
    const tableColumns = showPrevDay ? [...REPORT_COLUMNS, "Ngày TT T-1"] : REPORT_COLUMNS;

    let html = "<table><thead><tr>";
    tableColumns.forEach(header => {
        const isNameColumn = header === "Tên KH";
        html += `<th class="${isNameColumn ? "name-col" : ""}">${escapeHtml(header)}</th>`;
    });
    html += "</tr></thead><tbody>";

    data.forEach(row => {
        const urgent = normalizeProcessingLevel(row["Mức độ xử lý"]) === "CỰC KỲ GẤP";
        html += `<tr class="${urgent ? "urgent-row" : ""}">`;
        tableColumns.forEach(header => {
            let value = row[header] ?? "";
            if (header === "Ngày TT T-1") {
                const m = importMap[normImportHd(row.HD)];
                value = m && m.prevDay != null ? `Ngày ${m.prevDay}` : "";
            }
            if (header === "Tên KH") {
                html += `<td class="name-cell" title="${escapeHtml(String(value))}">${escapeHtml(value)}</td>`;
                return;
            }
            html += `<td>${escapeHtml(value)}</td>`;
        });
        html += "</tr>";
    });

    html += "</tbody></table>";
    container.innerHTML = html;
}

function setActiveReportTab(tabName) {
    document.querySelectorAll("[data-report-tab]").forEach(button => {
        const isActive = button.dataset.reportTab === tabName;
        button.classList.toggle("is-active", isActive);
        button.setAttribute("aria-selected", String(isActive));
        button.tabIndex = isActive ? 0 : -1;
    });

    document.querySelectorAll("[role='tabpanel']").forEach(panel => {
        panel.hidden = panel.id !== `panel-${tabName}`;
    });
}

function initializeReportTabs() {
    document.querySelectorAll("[data-report-tab]").forEach(button => {
        button.addEventListener("click", () => setActiveReportTab(button.dataset.reportTab));
    });
}

function initializeSearchableFilters() {
    if (!window.jQuery || !jQuery.fn.select2) return;

    jQuery("#filter-kv, #filter-dvkh").select2({
        width: "resolve",
        minimumResultsForSearch: 0,
        placeholder: "Chọn hoặc nhập để tìm",
        allowClear: true
    });
}

initializeReportTabs();
initializeSearchableFilters();
document.querySelectorAll("[data-collector-filter]").forEach(filter => {
    const toggle = filter.querySelector(".collector-filter-button");
    const menu = filter.querySelector(".collector-filter-menu");

    toggle.addEventListener("click", () => {
        const isOpen = menu.hidden;
        document.querySelectorAll(".collector-filter-menu").forEach(otherMenu => {
            otherMenu.hidden = true;
            otherMenu.closest("[data-collector-filter]")
                .querySelector(".collector-filter-button")
                .setAttribute("aria-expanded", "false");
        });
        menu.hidden = !isOpen;
        toggle.setAttribute("aria-expanded", String(isOpen));
        if (isOpen) filter.querySelector(".collector-filter-search").focus();
    });

    filter.querySelector(".collector-filter-search").addEventListener("input", () => {
        filterCollectorOptions(filter);
    });

    filter.addEventListener("change", event => {
        if (event.target.matches(".select-all-staff")) {
            selectedStaffs = event.target.checked ? [...collectorOptions] : [];
        } else if (event.target.matches(".collector-filter-options input[type='checkbox']")) {
            const staff = event.target.value;
            selectedStaffs = event.target.checked
                ? [...new Set([...selectedStaffs, staff])]
                : selectedStaffs.filter(selected => selected !== staff);
        } else {
            return;
        }

        syncCollectorControls();
        executeFilter();
    });
});
document.addEventListener("click", event => {
    const activeFilter = event.target instanceof Element
        ? event.target.closest("[data-collector-filter]")
        : null;
    document.querySelectorAll("[data-collector-filter]").forEach(filter => {
        if (filter === activeFilter) return;
        filter.querySelector(".collector-filter-menu").hidden = true;
        filter.querySelector(".collector-filter-button").setAttribute("aria-expanded", "false");
    });
});
document.addEventListener("keydown", event => {
    if (event.key !== "Escape") return;
    document.querySelectorAll(".collector-filter-menu").forEach(menu => {
        menu.hidden = true;
        menu.closest("[data-collector-filter]")
            .querySelector(".collector-filter-button")
            .setAttribute("aria-expanded", "false");
    });
});
document.getElementById("filter-hd").addEventListener("input", scheduleFilter);
document.getElementById("selectMucDoXuLy").addEventListener("change", executeFilter);
document.getElementById("filter-status").addEventListener("change", executeFilter);

document.getElementById("import-file").addEventListener("change", event => {
    const file = event.target.files && event.target.files[0];
    if (file) {
        handleImportFile(file).catch(err => alert("Lỗi import: " + err.message));
    }
    event.target.value = "";
});
renderImportStatus();
