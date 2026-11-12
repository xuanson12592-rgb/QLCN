const userAccounts = {
    "admin": { pass: "admin123", role: "admin", nvCode: "" },
    "HAILT60": { pass: "123456", role: "user", nvCode: "HAILT60" },
    "Thaonn11": { pass: "123456", role: "user", nvCode: "Thaonn11" }
};

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
        populateDropdowns(globalData);
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
        .toLocaleLowerCase("vi");
    filter.querySelectorAll(".collector-filter-options .collector-filter-option").forEach(option => {
        option.hidden = !option.dataset.staffName.toLocaleLowerCase("vi").includes(query);
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
    const { area, staff } = appliedFilters;
    const allCollectorsSelected = collectorOptions.length > 0
        && collectorOptions.every(collector => selectedStaffs.includes(collector));

    const filtered = globalData.filter(item => {
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

    data.forEach(item => {
        const collector = String(item["Nhân viên thu"] || "").trim() || "Chưa phân công";
        if (!groups.has(collector)) {
            groups.set(collector, {
                assigned: 0,
                paid: 0,
                cared: 0,
                notCared: 0,
            });
        }

        const summary = groups.get(collector);
        summary.assigned += 1;

        const paymentStatus = String(item["Tình Trạng"] || "").trim().toLocaleLowerCase("vi");
        const reconciliation = String(item["Kết quả đối soát"] || "").trim();
        if (reconciliation
            ? reconciliation === "Đã thanh toán"
            : /đã\s*(thanh toán|nộp tiền|thu|đóng)|đã\s*tt|paid|hoàn thành/.test(paymentStatus)) {
            summary.paid += 1;
        }

        const careStatus = String(item["Trạng thái chăm sóc"] || "").trim();
        if (careStatus === "Đã chăm sóc") summary.cared += 1;
        if (careStatus === "Chưa chăm sóc") summary.notCared += 1;

        const daysValue = item["Số Ngày TT muộn so với T-1"];
        if (daysValue !== "" && daysValue != null && Number.isFinite(Number(daysValue)) && Number(daysValue) > 0) {
            summary.late += 1;
        }
    });

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

    let html = "<table><thead><tr>";
    REPORT_COLUMNS.forEach(header => {
        const isNameColumn = header === "Tên KH";
        html += `<th class="${isNameColumn ? "name-col" : ""}">${escapeHtml(header)}</th>`;
    });
    html += "</tr></thead><tbody>";

    data.forEach(row => {
        const urgent = normalizeProcessingLevel(row["Mức độ xử lý"]) === "CỰC KỲ GẤP";
        html += `<tr class="${urgent ? "urgent-row" : ""}">`;
        REPORT_COLUMNS.forEach(header => {
            const value = row[header] ?? "";
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
