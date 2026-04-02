let parsedEmployees = [];
let sourceWorkbook = null;

const fileInput = document.getElementById("fileInput");
const generateBtn = document.getElementById("generateBtn");
const downloadBtn = document.getElementById("downloadBtn");
const printBtn = document.getElementById("printBtn");
const previewArea = document.getElementById("previewArea");
const statusBox = document.getElementById("status");
const monthInput = document.getElementById("monthInput");
const employeeFilter = document.getElementById("employeeFilter");
const inChargeNameInput = document.getElementById("inChargeName");
const regularHoursInput = document.getElementById("regularHours");
const saturdayHoursInput = document.getElementById("saturdayHours");
const paperSizeInput = document.getElementById("paperSize");

generateBtn.addEventListener("click", processFile);
downloadBtn.addEventListener("click", downloadExcel);
printBtn.addEventListener("click", () => window.print());

employeeFilter.addEventListener("change", renderPreview);
monthInput.addEventListener("input", renderPreview);
inChargeNameInput.addEventListener("input", renderPreview);
regularHoursInput.addEventListener("input", renderPreview);
saturdayHoursInput.addEventListener("input", renderPreview);
paperSizeInput.addEventListener("change", renderPreview);

function setStatus(msg, isError = false) {
    statusBox.textContent = msg;
    statusBox.style.color = isError ? "#b42318" : "#1d4f91";
}

function processFile() {
    const file = fileInput.files[0];
    if (!file) {
        setStatus("Please upload the ZKTeco Excel file first.", true);
        return;
    }

    const reader = new FileReader();
    reader.onload = function (e) {
        try {
            const data = new Uint8Array(e.target.result);
            sourceWorkbook = XLSX.read(data, { type: "array" });
            parsedEmployees = parseWorkbook(sourceWorkbook);

            if (!parsedEmployees.length) {
                setStatus("No Card Report employee data found in the workbook.", true);
                previewArea.innerHTML = "";
                employeeFilter.innerHTML = `<option value="">All Employees</option>`;
                return;
            }

            const detectedMonth = detectMonthFromWorkbook(sourceWorkbook);
            if (detectedMonth) monthInput.value = detectedMonth;

            populateEmployeeFilter(parsedEmployees);
            renderPreview();
            setStatus(`Generated ${parsedEmployees.length} employee DTR(s).`);
        } catch (error) {
            console.error(error);
            setStatus("Failed to read or parse the Excel file.", true);
        }
    };

    reader.readAsArrayBuffer(file);
}

function parseWorkbook(workbook) {
    const employees = [];

    for (const sheetName of workbook.SheetNames) {
        const sheet = workbook.Sheets[sheetName];
        const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });

        if (!isCardReportSheet(rows)) continue;
        employees.push(...parseCardReportSheet(rows));
    }

    return employees;
}

function isCardReportSheet(rows) {
    const hasCardReport = rows.some(row =>
        row.some(cell => normalizeCell(cell) === "Card Report")
    );
    const hasWeekDate = rows.some(row =>
        row.some(cell => normalizeCell(cell) === "WeekDate")
    );
    return hasCardReport && hasWeekDate;
}

function parseCardReportSheet(rows) {
    const employees = [];

    const headerRowIndex = rows.findIndex(row =>
        row.some(cell => normalizeCell(cell) === "Dept.")
    );
    const weekDateRowIndex = rows.findIndex(row =>
        row.some(cell => normalizeCell(cell) === "WeekDate")
    );

    if (headerRowIndex === -1 || weekDateRowIndex === -1) return employees;

    const headerRow = rows[headerRowIndex];
    const idRow = rows[headerRowIndex + 1] || [];

    for (let c = 0; c < headerRow.length; c++) {
        if (normalizeCell(headerRow[c]) === "Dept.") {
            let name = "";
            let id = "";
            let department = normalizeCell(headerRow[c + 1]);

            for (let x = c; x < c + 16 && x < headerRow.length; x++) {
                if (normalizeCell(headerRow[x]) === "Name") {
                    name = normalizeCell(headerRow[x + 1]);
                    break;
                }
            }

            for (let x = c; x < c + 16 && x < idRow.length; x++) {
                if (normalizeCell(idRow[x]) === "ID") {
                    id = normalizeCell(idRow[x + 1]);
                    break;
                }
            }

            if (name) {
                employees.push({
                    name,
                    employeeId: id,
                    department,
                    colStart: c,
                    colEnd: null,
                    logs: []
                });
            }
        }
    }

    employees.sort((a, b) => a.colStart - b.colStart);

    for (let i = 0; i < employees.length; i++) {
        const next = employees[i + 1];
        employees[i].colEnd = next ? next.colStart - 1 : rows[weekDateRowIndex].length - 1;
    }

    for (let r = weekDateRowIndex + 2; r < rows.length; r++) {
        const row = rows[r];

        employees.forEach(emp => {
            const block = row.slice(emp.colStart, emp.colEnd + 1).map(normalizeCell);
            const dayLabel = block[0];

            if (!/^\d{2}\s+[A-Z]{3}$/i.test(dayLabel)) return;

            const nonEmpty = block.slice(1).filter(Boolean);

            if (!nonEmpty.length) {
                emp.logs.push(buildEmptyLog(dayLabel));
                return;
            }

            if (nonEmpty[0].toLowerCase() === "absent") {
                emp.logs.push({
                    dayLabel,
                    dayNumber: parseDayNumber(dayLabel),
                    amArrival: "",
                    amDeparture: "",
                    pmArrival: "",
                    pmDeparture: "",
                    undertimeHours: "",
                    undertimeMinutes: "",
                    status: "Absent"
                });
                return;
            }

            const times = nonEmpty.filter(isTimeString);
            const amArrival = times[0] || "";
            const amDeparture = times[1] || "";
            const pmArrival = times[2] || "";
            const pmDeparture = times[3] || "";

            const undertimeTotalMinutes = computeUndertimeTotalMinutes(pmDeparture, "17:00");

            emp.logs.push({
                dayLabel,
                dayNumber: parseDayNumber(dayLabel),
                amArrival,
                amDeparture,
                pmArrival,
                pmDeparture,
                undertimeHours: undertimeTotalMinutes ? Math.floor(undertimeTotalMinutes / 60) : "",
                undertimeMinutes: undertimeTotalMinutes ? undertimeTotalMinutes % 60 : "",
                status: classifyStatus(dayLabel, amArrival, amDeparture, pmArrival, pmDeparture)
            });
        });
    }

    const monthValue = detectMonthFromRows(rows);
    if (monthValue) {
        const [year, month] = monthValue.split("-").map(Number);
        const daysInMonth = new Date(year, month, 0).getDate();

        employees.forEach(emp => {
            const map = new Map(emp.logs.map(log => [log.dayNumber, log]));
            const completed = [];

            for (let day = 1; day <= daysInMonth; day++) {
                if (map.has(day)) completed.push(map.get(day));
                else completed.push(buildEmptyLog(formatDayLabel(year, month, day)));
            }

            emp.logs = completed;
        });
    }

    return employees;
}

function populateEmployeeFilter(employees) {
    employeeFilter.innerHTML = `<option value="">All Employees</option>`;
    employees.forEach((emp, index) => {
        const option = document.createElement("option");
        option.value = String(index);
        option.textContent = emp.name;
        employeeFilter.appendChild(option);
    });
}

function getPreviewConfig() {
    return {
        titleTop: "9px",
        titleMain: "12px",
        divider: "8px",
        metaLabel: "7px",
        metaLine: "7.2px",
        metaNote: "5.5px",
        hours: "6.5px",
        table: "6.7px",
        subhead: "6.1px",
        certification: "5.8px",
        verified: "6.2px",
        incharge: "6.5px",
        inchargeLabel: "5.8px",
        dailyRowHeight: "10px",
        lineHeight: "1"
    };
}

function renderPreview() {
    if (!parsedEmployees.length) {
        previewArea.innerHTML = "";
        return;
    }

    const selected = employeeFilter.value;
    const employeesToShow = selected === ""
        ? parsedEmployees
        : [parsedEmployees[Number(selected)]].filter(Boolean);

    const monthText = formatMonthDisplay(monthInput.value);
    const inChargeName = inChargeNameInput.value.trim();
    const regularHours = regularHoursInput.value.trim();
    const saturdayHours = saturdayHoursInput.value.trim();
    const paperSize = paperSizeInput.value;
    const previewConfig = getPreviewConfig();

    previewArea.innerHTML = employeesToShow.map(emp => `
        <div class="form48-page ${paperSize}">
            <div class="form48-row">
                ${buildForm48CopyHTML(emp, monthText, regularHours, saturdayHours, inChargeName, previewConfig)}
                <div class="cut-gap"></div>
                ${buildForm48CopyHTML(emp, monthText, regularHours, saturdayHours, inChargeName, previewConfig)}
                <div class="cut-gap"></div>
                ${buildForm48CopyHTML(emp, monthText, regularHours, saturdayHours, inChargeName, previewConfig)}
            </div>
        </div>
    `).join("");
}

function buildForm48CopyHTML(emp, monthText, regularHours, saturdayHours, inChargeName, cfg) {
    const totals = summarizeUndertime(emp.logs);

    return `
        <div class="form48-copy">
            <div class="copy-top">
                <div class="form48-header">
                    <div class="form48-title-top" style="font-size:${cfg.titleTop}">Civil Service Form No. 48</div>
                    <div class="form48-title-main" style="font-size:${cfg.titleMain}">DAILY TIME RECORD</div>
                    <div class="form48-divider" style="font-size:${cfg.divider}">-----o0o-----</div>
                </div>

                <div class="meta-label" style="font-size:${cfg.metaLabel}">&nbsp;</div>
                <div class="meta-line" style="font-size:${cfg.metaLine}">${escapeHtml(emp.name)}</div>
                <div class="meta-note" style="font-size:${cfg.metaNote}">(Name)</div>

                <div class="meta-label" style="font-size:${cfg.metaLabel}">For the month of</div>
                <div class="meta-line" style="font-size:${cfg.metaLine}">${escapeHtml(monthText)}</div>

                <div class="hours-block" style="font-size:${cfg.hours}">
                    <div>Official hours for arrival and departure</div>
                    <div class="hours-row">
                        <div>Regular days</div>
                        <div class="hours-value">${escapeHtml(regularHours)}</div>
                    </div>
                    <div class="hours-row">
                        <div>Saturdays</div>
                        <div class="hours-value">${escapeHtml(saturdayHours)}</div>
                    </div>
                </div>

                <table class="form48-table" style="font-size:${cfg.table}">
                    <thead>
                        <tr>
                            <th rowspan="2" style="width:10%;">Day</th>
                            <th colspan="2" style="width:30%;">A.M.</th>
                            <th colspan="2" style="width:30%;">P.M.</th>
                            <th colspan="2" style="width:30%;">Undertime</th>
                        </tr>
                        <tr class="subhead" style="font-size:${cfg.subhead}">
                            <th>Arrival</th>
                            <th>Depar-ture</th>
                            <th>Arrival</th>
                            <th>Depar-ture</th>
                            <th>Hours</th>
                            <th>Min-utes</th>
                        </tr>
                    </thead>
                    <tbody>
                        ${emp.logs.map(log => `
                            <tr style="height:${cfg.dailyRowHeight}; line-height:${cfg.lineHeight}">
                                <td>${log.dayNumber || ""}</td>
                                <td>${displayTime(log, "amArrival")}</td>
                                <td>${displayTime(log, "amDeparture")}</td>
                                <td>${displayTime(log, "pmArrival")}</td>
                                <td>${displayTime(log, "pmDeparture")}</td>
                                <td>${displayUndertime(log, "hours")}</td>
                                <td>${displayUndertime(log, "minutes")}</td>
                            </tr>
                        `).join("")}
                        <tr class="total-row" style="height:${cfg.dailyRowHeight}; line-height:${cfg.lineHeight}">
                            <td>Total</td>
                            <td></td>
                            <td></td>
                            <td></td>
                            <td></td>
                            <td>${totals.hours}</td>
                            <td>${totals.minutes}</td>
                        </tr>
                    </tbody>
                </table>
            </div>

            <div class="copy-bottom">
                <div class="certification" style="font-size:${cfg.certification}">
                    I certify on my honor that the above is a true and correct report of the hours of work performed,
                    record of which was made daily at the time of arrival and departure from office.
                </div>

                <div class="verified" style="font-size:${cfg.verified}">
                    <strong>VERIFIED as to the prescribed office hours:</strong>
                    <div class="incharge-wrap">
                        <div class="incharge-line" style="font-size:${cfg.incharge}">${escapeHtml(inChargeName)}</div>
                        <div class="incharge-label" style="font-size:${cfg.inchargeLabel}">In Charge</div>
                    </div>
                </div>
            </div>
        </div>
    `;
}

function displayTime(log, key) {
    if (log.status === "Absent") return key === "amArrival" ? "ABSENT" : "";
    return escapeHtml(log[key] || "");
}

function displayUndertime(log, type) {
    if (log.status === "Absent") return "";
    return type === "hours" ? (log.undertimeHours ?? "") : (log.undertimeMinutes ?? "");
}

function summarizeUndertime(logs) {
    let totalMinutes = 0;
    logs.forEach(log => {
        const h = Number(log.undertimeHours || 0);
        const m = Number(log.undertimeMinutes || 0);
        totalMinutes += h * 60 + m;
    });
    return {
        hours: totalMinutes ? Math.floor(totalMinutes / 60) : "",
        minutes: totalMinutes ? totalMinutes % 60 : ""
    };
}

function detectMonthFromWorkbook(workbook) {
    for (const sheetName of workbook.SheetNames) {
        const sheet = workbook.Sheets[sheetName];
        const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
        const found = detectMonthFromRows(rows);
        if (found) return found;
    }
    return "";
}

function detectMonthFromRows(rows) {
    for (const row of rows) {
        for (const cell of row) {
            const text = normalizeCell(cell);
            const match = text.match(/(\d{4})-(\d{2})-\d{2}\s*~\s*(\d{4})-(\d{2})-\d{2}/);
            if (match) return `${match[1]}-${match[2]}`;
        }
    }
    return "";
}

function buildEmptyLog(dayLabel) {
    return {
        dayLabel,
        dayNumber: parseDayNumber(dayLabel),
        amArrival: "",
        amDeparture: "",
        pmArrival: "",
        pmDeparture: "",
        undertimeHours: "",
        undertimeMinutes: "",
        status: ""
    };
}

function parseDayNumber(dayLabel) {
    const match = String(dayLabel).match(/^(\d{2})/);
    return match ? Number(match[1]) : null;
}

function normalizeCell(cell) {
    if (cell === null || cell === undefined) return "";
    return String(cell).trim();
}

function isTimeString(value) {
    return /^\d{1,2}:\d{2}$/.test(String(value).trim());
}

function toMinutes(timeStr) {
    if (!isTimeString(timeStr)) return null;
    const [h, m] = String(timeStr).split(":").map(Number);
    return h * 60 + m;
}

function computeUndertimeTotalMinutes(timeOut, standardOut = "17:00") {
    const actual = toMinutes(timeOut);
    const standard = toMinutes(standardOut);
    if (actual === null || standard === null) return 0;
    return actual < standard ? standard - actual : 0;
}

function classifyStatus(dayLabel, amArrival, amDeparture, pmArrival, pmDeparture) {
    const dayName = dayLabel.split(" ")[1]?.toUpperCase() || "";
    if (dayName === "SAT" || dayName === "SUN") return "Weekend";
    if (!amArrival && !amDeparture && !pmArrival && !pmDeparture) return "";
    return "Present";
}

function formatDayLabel(year, month, day) {
    const date = new Date(year, month - 1, day);
    const weekday = date.toLocaleDateString("en-US", { weekday: "short" }).toUpperCase();
    return `${String(day).padStart(2, "0")} ${weekday}`;
}

function formatMonthDisplay(value) {
    if (!value) return "";
    const [year, month] = value.split("-").map(Number);
    const date = new Date(year, month - 1, 1);
    return date.toLocaleDateString("en-US", {
        month: "long",
        year: "numeric"
    });
}

function escapeHtml(str) {
    return String(str ?? "")
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;");
}

function safeSheetName(name) {
    return name.replace(/[\\/*?:[\]]/g, "").slice(0, 31) || "DTR";
}

async function downloadExcel() {
    if (!parsedEmployees.length) {
        setStatus("No generated DTR to export yet.", true);
        return;
    }

    try {
        const workbook = new ExcelJS.Workbook();
        workbook.creator = "ChatGPT";
        workbook.created = new Date();

        const selected = employeeFilter.value;
        const employeesToExport = selected === ""
            ? parsedEmployees
            : [parsedEmployees[Number(selected)]].filter(Boolean);

        const monthText = formatMonthDisplay(monthInput.value);
        const regularHours = regularHoursInput.value.trim();
        const saturdayHours = saturdayHoursInput.value.trim();
        const inChargeName = inChargeNameInput.value.trim();
        const paperSize = paperSizeInput.value;

        for (const emp of employeesToExport) {
            const ws = workbook.addWorksheet(safeSheetName(emp.name), {
                pageSetup: {
                    paperSize: paperSize === "folio" ? 14 : 9,
                    orientation: "landscape",
                    fitToPage: true,
                    fitToWidth: 1,
                    fitToHeight: 1,
                    margins: {
                        left: 0.05,
                        right: 0.05,
                        top: 0.05,
                        bottom: 0.05,
                        header: 0.03,
                        footer: 0.03
                    }
                }
            });

            buildExcelHorizontalTriplicate(ws, emp, {
                monthText,
                regularHours,
                saturdayHours,
                inChargeName
            });
        }

        const buffer = await workbook.xlsx.writeBuffer();
        saveAs(new Blob([buffer]), "CS_Form48_3Copies_Landscape_FullHeight.xlsx");
        setStatus("Landscape Excel downloaded successfully.");
    } catch (error) {
        console.error(error);
        setStatus("Failed to generate the Excel file.", true);
    }
}

function buildExcelHorizontalTriplicate(ws, emp, options) {
    const { monthText, regularHours, saturdayHours, inChargeName } = options;

    const leftStart = 1;
    const midStart = 10;
    const rightStart = 19;

    ws.columns = [
        { width: 6.8 }, { width: 8.3 }, { width: 8.3 }, { width: 8.3 }, { width: 8.3 }, { width: 6.4 }, { width: 6.4 }, { width: 1.3 }, { width: 1.3 },
        { width: 6.8 }, { width: 8.3 }, { width: 8.3 }, { width: 8.3 }, { width: 8.3 }, { width: 6.4 }, { width: 6.4 }, { width: 1.3 }, { width: 1.3 },
        { width: 6.8 }, { width: 8.3 }, { width: 8.3 }, { width: 8.3 }, { width: 8.3 }, { width: 6.4 }, { width: 6.4 }
    ];

    const center = { vertical: "middle", horizontal: "center", wrapText: true };
    const left = { vertical: "middle", horizontal: "left", wrapText: true };
    const thinBorder = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" }
    };
    const texturedFill = {
        type: "pattern",
        pattern: "darkVertical",
        fgColor: { argb: "FFF1F1F1" },
        bgColor: { argb: "FFFFFFFF" }
    };
    const dashedGap = {
        left: { style: "dashed" },
        right: { style: "dashed" }
    };

    buildSingleHorizontalCopy(ws, leftStart, emp, monthText, regularHours, saturdayHours, inChargeName, center, left, thinBorder);
    buildSingleHorizontalCopy(ws, midStart, emp, monthText, regularHours, saturdayHours, inChargeName, center, left, thinBorder);
    buildSingleHorizontalCopy(ws, rightStart, emp, monthText, regularHours, saturdayHours, inChargeName, center, left, thinBorder);

    ws.getRow(1).height = 12;
    ws.getRow(2).height = 13;
    ws.getRow(3).height = 10;
    ws.getRow(4).height = 7;
    ws.getRow(5).height = 13;
    ws.getRow(6).height = 9;
    ws.getRow(7).height = 7;
    ws.getRow(8).height = 12;
    ws.getRow(9).height = 7;
    ws.getRow(10).height = 10;
    ws.getRow(11).height = 11;
    ws.getRow(12).height = 11;
    ws.getRow(13).height = 7;
    ws.getRow(14).height = 12;
    ws.getRow(15).height = 13;

    for (let r = 16; r <= 46; r++) {
        ws.getRow(r).height = 13.5;
    }

    ws.getRow(47).height = 12;
    ws.getRow(48).height = 12;
    ws.getRow(49).height = 12;

    for (let r = 1; r <= 49; r++) {
        [8, 9, 17, 18].forEach(col => {
            ws.getCell(r, col).fill = texturedFill;
            ws.getCell(r, col).border = dashedGap;
        });
    }
}

function buildSingleHorizontalCopy(ws, startCol, emp, monthText, regularHours, saturdayHours, inChargeName, center, left, thinBorder) {
    const totals = summarizeUndertime(emp.logs);
    const col = i => ws.getColumn(startCol + i - 1).letter;
    const A = col(1), B = col(2), C = col(3), D = col(4), E = col(5), F = col(6), G = col(7);

    ws.mergeCells(`${A}1:${G}1`);
    ws.getCell(`${A}1`).value = "Civil Service Form No. 48";
    ws.getCell(`${A}1`).font = { bold: true, size: 9.4 };
    ws.getCell(`${A}1`).alignment = center;

    ws.mergeCells(`${A}2:${G}2`);
    ws.getCell(`${A}2`).value = "DAILY TIME RECORD";
    ws.getCell(`${A}2`).font = { bold: true, size: 10.9 };
    ws.getCell(`${A}2`).alignment = center;

    ws.mergeCells(`${A}3:${G}3`);
    ws.getCell(`${A}3`).value = "-----o0o-----";
    ws.getCell(`${A}3`).font = { size: 7.9 };
    ws.getCell(`${A}3`).alignment = center;

    ws.mergeCells(`${B}5:${F}5`);
    ws.getCell(`${B}5`).value = emp.name;
    ws.getCell(`${B}5`).font = { bold: true, size: 8.1 };
    ws.getCell(`${B}5`).alignment = center;
    ws.getCell(`${B}5`).border = { bottom: { style: "thin" } };

    ws.mergeCells(`${B}6:${F}6`);
    ws.getCell(`${B}6`).value = "(Name)";
    ws.getCell(`${B}6`).font = { size: 6.5 };
    ws.getCell(`${B}6`).alignment = center;

    ws.getCell(`${A}8`).value = "For the month of";
    ws.getCell(`${A}8`).font = { size: 7 };

    ws.mergeCells(`${C}8:${G}8`);
    ws.getCell(`${C}8`).value = monthText;
    ws.getCell(`${C}8`).font = { bold: true, size: 7.7 };
    ws.getCell(`${C}8`).alignment = center;
    ws.getCell(`${C}8`).border = { bottom: { style: "thin" } };

    ws.mergeCells(`${A}10:${G}10`);
    ws.getCell(`${A}10`).value = "Official hours for arrival and departure";
    ws.getCell(`${A}10`).font = { size: 7 };
    ws.getCell(`${A}10`).alignment = left;

    ws.getCell(`${A}11`).value = "Regular days";
    ws.getCell(`${A}11`).font = { size: 7 };
    ws.mergeCells(`${C}11:${G}11`);
    ws.getCell(`${C}11`).value = regularHours;
    ws.getCell(`${C}11`).font = { size: 7 };
    ws.getCell(`${C}11`).border = { bottom: { style: "thin" } };

    ws.getCell(`${A}12`).value = "Saturdays";
    ws.getCell(`${A}12`).font = { size: 7 };
    ws.mergeCells(`${C}12:${G}12`);
    ws.getCell(`${C}12`).value = saturdayHours;
    ws.getCell(`${C}12`).font = { size: 7 };
    ws.getCell(`${C}12`).border = { bottom: { style: "thin" } };

    ws.mergeCells(`${A}14:${A}15`);
    ws.getCell(`${A}14`).value = "Day";
    ws.getCell(`${A}14`).font = { bold: true, size: 7 };
    ws.getCell(`${A}14`).alignment = center;

    ws.mergeCells(`${B}14:${C}14`);
    ws.getCell(`${B}14`).value = "A.M.";
    ws.getCell(`${B}14`).font = { bold: true, size: 7 };
    ws.getCell(`${B}14`).alignment = center;

    ws.mergeCells(`${D}14:${E}14`);
    ws.getCell(`${D}14`).value = "P.M.";
    ws.getCell(`${D}14`).font = { bold: true, size: 7 };
    ws.getCell(`${D}14`).alignment = center;

    ws.mergeCells(`${F}14:${G}14`);
    ws.getCell(`${F}14`).value = "Undertime";
    ws.getCell(`${F}14`).font = { bold: true, size: 7 };
    ws.getCell(`${F}14`).alignment = center;

    ws.getCell(`${B}15`).value = "Arrival";
    ws.getCell(`${C}15`).value = "Departure";
    ws.getCell(`${D}15`).value = "Arrival";
    ws.getCell(`${E}15`).value = "Departure";
    ws.getCell(`${F}15`).value = "Hours";
    ws.getCell(`${G}15`).value = "Minutes";

    [B, C, D, E, F, G].forEach(letter => {
        ws.getCell(`${letter}15`).font = { bold: true, size: 6.5 };
        ws.getCell(`${letter}15`).alignment = center;
    });

    let rowNum = 16;

    for (const log of emp.logs) {
        ws.getCell(`${A}${rowNum}`).value = log.dayNumber || "";
        ws.getCell(`${B}${rowNum}`).value = log.status === "Absent" ? "ABSENT" : (log.amArrival || "");
        ws.getCell(`${C}${rowNum}`).value = log.status === "Absent" ? "" : (log.amDeparture || "");
        ws.getCell(`${D}${rowNum}`).value = log.status === "Absent" ? "" : (log.pmArrival || "");
        ws.getCell(`${E}${rowNum}`).value = log.status === "Absent" ? "" : (log.pmDeparture || "");
        ws.getCell(`${F}${rowNum}`).value = log.status === "Absent" ? "" : (log.undertimeHours || "");
        ws.getCell(`${G}${rowNum}`).value = log.status === "Absent" ? "" : (log.undertimeMinutes || "");
        rowNum++;
    }

    ws.getCell(`${A}${rowNum}`).value = "Total";
    ws.getCell(`${F}${rowNum}`).value = totals.hours;
    ws.getCell(`${G}${rowNum}`).value = totals.minutes;

    for (let r = 14; r <= rowNum; r++) {
        [A, B, C, D, E, F, G].forEach(letter => {
            ws.getCell(`${letter}${r}`).border = thinBorder;
            ws.getCell(`${letter}${r}`).alignment = center;
            if (!ws.getCell(`${letter}${r}`).font) {
                ws.getCell(`${letter}${r}`).font = { size: 6.5 };
            }
        });
    }
const gap = 2; // number of blank rows above

ws.mergeCells(`${A}${rowNum + gap}:${G}${rowNum + gap}`);
ws.getCell(`${A}${rowNum + gap}`).value = "I certify on my honor that the above is a true and correct report of the hours of work performed,";
ws.getCell(`${A}${rowNum + gap}`).font = { size: 7.1 };
ws.getCell(`${A}${rowNum + gap}`).alignment = left;

ws.mergeCells(`${A}${rowNum + gap + 1}:${G}${rowNum + gap + 1}`);
ws.getCell(`${A}${rowNum + gap + 1}`).value = "record of which was made daily at the time of arrival and departure from office.";
ws.getCell(`${A}${rowNum + gap + 1}`).font = { size: 7.1 };
ws.getCell(`${A}${rowNum + gap + 1}`).alignment = left;

ws.mergeCells(`${A}${rowNum + gap + 3}:${G}${rowNum + gap + 3}`);
ws.getCell(`${A}${rowNum + gap + 3}`).value = "VERIFIED as to the prescribed office hours:";
ws.getCell(`${A}${rowNum + gap + 3}`).font = { bold: true, size: 7.1 };
ws.getCell(`${A}${rowNum + gap + 3}`).alignment = left;

ws.mergeCells(`${D}${rowNum + gap + 5}:${G}${rowNum + gap + 5}`);
ws.getCell(`${D}${rowNum + gap + 5}`).value = inChargeName;
ws.getCell(`${D}${rowNum + gap + 5}`).font = { bold: true, size: 7.8 };
ws.getCell(`${D}${rowNum + gap + 5}`).alignment = center;
ws.getCell(`${D}${rowNum + gap + 5}`).border = { bottom: { style: "thin" } };

ws.mergeCells(`${D}${rowNum + gap + 6}:${G}${rowNum + gap + 6}`);
ws.getCell(`${D}${rowNum + gap + 6}`).value = "In Charge";
ws.getCell(`${D}${rowNum + gap + 6}`).font = { size: 7.1 };
ws.getCell(`${D}${rowNum + gap + 6}`).alignment = center;




}