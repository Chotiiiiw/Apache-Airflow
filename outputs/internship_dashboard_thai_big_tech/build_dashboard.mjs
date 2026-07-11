import fs from "node:fs/promises";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const outputDir = "/Users/chotiwatbanjongpru/Desktop/ Airflow_tutorial/outputs/internship_dashboard_thai_big_tech";
const outputPath = `${outputDir}/thai_big_tech_internship_dashboard.xlsx`;

const workbook = Workbook.create();

const dashboard = workbook.worksheets.add("Dashboard");
const apps = workbook.worksheets.add("Applications");
const companies = workbook.worksheets.add("Company Targets");
const contacts = workbook.worksheets.add("Contacts");
const lookups = workbook.worksheets.add("Lookups");
const readme = workbook.worksheets.add("README");

const palette = {
  ink: "#1F2933",
  muted: "#667085",
  bg: "#F6F8FA",
  panel: "#FFFFFF",
  line: "#D9E2EC",
  navy: "#243B53",
  teal: "#0F766E",
  green: "#2F855A",
  amber: "#B7791F",
  red: "#C53030",
  blue: "#2563EB",
  purple: "#6D28D9",
  softTeal: "#DDF7F3",
  softBlue: "#EAF2FF",
  softAmber: "#FFF4D6",
  softRed: "#FCE8E8",
  softGreen: "#E6F4EA",
};

const statusList = [
  "Saved",
  "Preparing",
  "Applied",
  "OA / Test",
  "Interview",
  "Offer",
  "Rejected",
  "Withdrawn",
];
const priorityList = ["High", "Medium", "Low"];
const sourceList = ["Careers Page", "LinkedIn", "JobsDB", "Referral", "Campus", "Event", "Other"];
const workModeList = ["Bangkok", "Hybrid", "Remote", "On-site", "Thailand"];
const roleTypeList = ["Software Engineer", "Data", "Product", "UX/UI", "Business", "Cybersecurity", "AI/ML", "Cloud/DevOps"];

const appHeaders = [
  "Company",
  "Role",
  "Role Type",
  "Location / Mode",
  "Source",
  "Status",
  "Priority",
  "Deadline",
  "Applied Date",
  "Next Follow-up",
  "Fit Score",
  "Referral",
  "Link",
  "Notes",
];

const appRows = [
  ["LINE MAN Wongnai", "Software Engineer Intern", "Software Engineer", "Bangkok", "Careers Page", "Applied", "High", new Date("2026-07-25"), new Date("2026-07-05"), new Date("2026-07-15"), 9, "Ploy - Alumni", "https://lmwn.com/careers", "Backend team, food delivery platform"],
  ["Agoda Thailand", "Data Analyst Intern", "Data", "Bangkok", "LinkedIn", "Interview", "High", new Date("2026-07-20"), new Date("2026-07-01"), new Date("2026-07-13"), 8, "None", "https://careersatagoda.com", "Prepare SQL case and product metrics"],
  ["Ascend Money / TrueMoney", "Product Intern", "Product", "Hybrid", "Careers Page", "Preparing", "High", new Date("2026-07-31"), null, new Date("2026-07-14"), 8, "Somchai - Friend", "https://www.ascendcorp.com/careers", "Customize CV for payments and fintech"],
  ["KBTG", "AI Engineer Intern", "AI/ML", "Bangkok", "Campus", "OA / Test", "High", new Date("2026-07-18"), new Date("2026-07-03"), new Date("2026-07-12"), 9, "University event", "https://www.kbtg.tech/careers", "Review ML basics and banking use cases"],
  ["SCB 10X", "Venture Tech Intern", "Business", "Hybrid", "Referral", "Saved", "Medium", new Date("2026-08-10"), null, new Date("2026-07-18"), 7, "Mint - LinkedIn", "https://www.scb10x.com/careers", "Needs tailored cover letter"],
  ["Bitkub", "Cybersecurity Intern", "Cybersecurity", "Bangkok", "JobsDB", "Applied", "Medium", new Date("2026-07-28"), new Date("2026-07-08"), new Date("2026-07-16"), 7, "None", "https://www.bitkub.com/career", "Highlight CTF and security coursework"],
  ["Sertis", "Machine Learning Intern", "AI/ML", "Bangkok", "Careers Page", "Preparing", "Medium", new Date("2026-08-05"), null, new Date("2026-07-17"), 8, "None", "https://www.sertiscorp.com/careers", "Build mini portfolio note"],
  ["Central Tech", "Cloud Engineer Intern", "Cloud/DevOps", "Hybrid", "LinkedIn", "Saved", "Low", new Date("2026-08-12"), null, new Date("2026-07-20"), 6, "None", "https://www.central.tech/careers", "Check tech stack before applying"],
  ["CP Axtra Digital", "UX Research Intern", "UX/UI", "Bangkok", "Event", "Applied", "Medium", new Date("2026-07-30"), new Date("2026-07-06"), new Date("2026-07-19"), 7, "Recruiter event", "https://www.cpaxtra.com/en/careers", "Retail tech research angle"],
  ["True Digital Group", "Software Engineer Intern", "Software Engineer", "Bangkok", "Careers Page", "Rejected", "Low", new Date("2026-07-15"), new Date("2026-06-25"), null, 6, "None", "https://www.truedigital.com/careers", "Keep contact warm for next cycle"],
];

const companyHeaders = ["Company", "Focus Area", "Target Role", "Priority", "Career Link", "Why it fits", "Last Checked", "Next Action"];
const companyRows = [
  ["LINE MAN Wongnai", "Food delivery, marketplace, payments", "Software Engineer / Product", "High", "https://lmwn.com/careers", "High-scale Thai consumer platform", new Date("2026-07-10"), "Track backend and data roles"],
  ["Agoda Thailand", "Travel tech, data platform", "Data / Software Engineer", "High", "https://careersatagoda.com", "Strong analytics and engineering culture in Bangkok", new Date("2026-07-10"), "Prepare SQL and product cases"],
  ["KBTG", "Banking tech, AI, platform engineering", "AI/ML / Software Engineer", "High", "https://www.kbtg.tech/careers", "Large Thai tech org with applied AI work", new Date("2026-07-10"), "Follow campus recruiting"],
  ["Ascend Money / TrueMoney", "Fintech, wallet, payments", "Product / Data / Software Engineer", "High", "https://www.ascendcorp.com/careers", "Fintech exposure and regional scale", new Date("2026-07-10"), "Ask referral contact"],
  ["SCB 10X", "Venture, fintech, digital assets", "Product / Venture Tech", "Medium", "https://www.scb10x.com/careers", "Good fit for product plus strategy track", new Date("2026-07-10"), "Tailor story to innovation"],
  ["Bitkub", "Digital assets, exchange platform", "Cybersecurity / Data", "Medium", "https://www.bitkub.com/career", "Security and financial platform experience", new Date("2026-07-10"), "Review compliance basics"],
  ["Sertis", "AI, machine learning, enterprise data", "AI/ML / Data", "Medium", "https://www.sertiscorp.com/careers", "Strong applied machine learning target", new Date("2026-07-10"), "Prepare ML portfolio"],
  ["Central Tech", "Retail tech, ecommerce, cloud", "Cloud/DevOps / Software Engineer", "Medium", "https://www.central.tech/careers", "Retail platform and cloud exposure", new Date("2026-07-10"), "Check current internship openings"],
  ["CP Axtra Digital", "Retail, wholesale, digital transformation", "UX/UI / Product / Data", "Medium", "https://www.cpaxtra.com/en/careers", "Useful for retail tech and operations", new Date("2026-07-10"), "Map roles to portfolio"],
  ["True Digital Group", "Telecom digital products, platforms", "Software Engineer / Product", "Medium", "https://www.truedigital.com/careers", "Large ecosystem and consumer product work", new Date("2026-07-10"), "Monitor next intake"],
];

const contactHeaders = ["Name", "Company", "Role / Relationship", "Channel", "Contacted Date", "Follow-up Date", "Status", "Notes"];
const contactRows = [
  ["Ploy", "LINE MAN Wongnai", "Alumni engineer", "LinkedIn", new Date("2026-07-02"), new Date("2026-07-16"), "Warm", "Asked about backend intern team"],
  ["Somchai", "Ascend Money / TrueMoney", "Friend of friend", "LINE", new Date("2026-07-08"), new Date("2026-07-14"), "Warm", "Can review CV"],
  ["Mint", "SCB 10X", "LinkedIn product contact", "LinkedIn", new Date("2026-07-09"), new Date("2026-07-18"), "New", "Send short intro after profile update"],
  ["Recruiter event", "CP Axtra Digital", "Campus event contact", "Email", new Date("2026-07-06"), new Date("2026-07-19"), "Warm", "Follow up with portfolio link"],
];

function setWidths(sheet, widths) {
  widths.forEach((width, idx) => {
    sheet.getRangeByIndexes(0, idx, 1, 1).format.columnWidth = width;
  });
}

function styleTitle(sheet, range, title, subtitle) {
  range.merge();
  range.values = [[title]];
  range.format.fill.color = palette.navy;
  range.format.font.color = "#FFFFFF";
  range.format.font.bold = true;
  range.format.font.size = 18;
  range.format.rowHeight = 34;
  const sub = sheet.getRange("A2:N2");
  sub.merge();
  sub.values = [[subtitle]];
  sub.format.fill.color = palette.navy;
  sub.format.font.color = "#DDE7F0";
  sub.format.font.size = 10;
}

function styleHeader(range) {
  range.format.fill.color = palette.teal;
  range.format.font.color = "#FFFFFF";
  range.format.font.bold = true;
  range.format.wrapText = true;
  range.format.borders = { preset: "outside", style: "thin", color: palette.teal };
}

function styleTable(range) {
  range.format.borders = {
    insideHorizontal: { style: "thin", color: "#E5E7EB" },
    insideVertical: { style: "thin", color: "#EEF2F7" },
    bottom: { style: "thin", color: palette.line },
  };
  range.format.font.color = palette.ink;
  range.format.font.size = 10;
}

function writeMatrix(sheet, startRow, startCol, matrix) {
  sheet.getRangeByIndexes(startRow, startCol, matrix.length, matrix[0].length).values = matrix;
}

function addListValidation(sheet, range, values) {
  sheet.getRange(range).dataValidation = {
    rule: { type: "list", values },
  };
}

for (const sheet of [dashboard, apps, companies, contacts, lookups, readme]) {
  sheet.showGridLines = false;
}

// Lookups
lookups.getRange("A1:F1").values = [["Status", "Priority", "Source", "Location / Mode", "Role Type", "Status Order"]];
styleHeader(lookups.getRange("A1:F1"));
const maxRows = Math.max(statusList.length, priorityList.length, sourceList.length, workModeList.length, roleTypeList.length);
const lookupRows = Array.from({ length: maxRows }, (_, i) => [
  statusList[i] ?? null,
  priorityList[i] ?? null,
  sourceList[i] ?? null,
  workModeList[i] ?? null,
  roleTypeList[i] ?? null,
  statusList[i] ? i + 1 : null,
]);
writeMatrix(lookups, 1, 0, lookupRows);
styleTable(lookups.getRange("A2:F10"));
setWidths(lookups, [18, 14, 18, 18, 20, 14]);

// Applications
styleTitle(apps, apps.getRange("A1:N1"), "Thai Big Tech Internship Tracker", "Update rows 6-105. Dashboard formulas read from this table.");
apps.getRange("A5:N5").values = [appHeaders];
styleHeader(apps.getRange("A5:N5"));
writeMatrix(apps, 5, 0, appRows);
styleTable(apps.getRange("A6:N105"));
setWidths(apps, [22, 24, 18, 16, 16, 15, 12, 13, 13, 14, 10, 20, 34, 36]);
apps.getRange("H6:J105").setNumberFormat("yyyy-mm-dd");
apps.getRange("K6:K105").setNumberFormat("0");
apps.getRange("M6:M105").format.font.color = palette.blue;
apps.getRange("N6:N105").format.wrapText = true;
apps.freezePanes.freezeRows(5);
apps.freezePanes.freezeColumns(2);
addListValidation(apps, "C6:C105", roleTypeList);
addListValidation(apps, "D6:D105", workModeList);
addListValidation(apps, "E6:E105", sourceList);
addListValidation(apps, "F6:F105", statusList);
addListValidation(apps, "G6:G105", priorityList);
apps.getRange("K6:K105").dataValidation = { rule: { type: "whole", operator: "between", formula1: 1, formula2: 10 } };
apps.getRange("A4:N4").merge();
apps.getRange("A4:N4").values = [["Tip: keep one row per role. Use status and next follow-up to drive the dashboard."]];
apps.getRange("A4:N4").format.fill.color = palette.softBlue;
apps.getRange("A4:N4").format.font.color = palette.navy;

// Company Targets
styleTitle(companies, companies.getRange("A1:H1"), "Company Targets", "A practical shortlist for internship hunting in Thai big tech and tech-forward companies.");
companies.getRange("A5:H5").values = [companyHeaders];
styleHeader(companies.getRange("A5:H5"));
writeMatrix(companies, 5, 0, companyRows);
styleTable(companies.getRange("A6:H40"));
setWidths(companies, [24, 32, 28, 12, 36, 42, 13, 28]);
companies.getRange("G6:G40").setNumberFormat("yyyy-mm-dd");
companies.getRange("E6:E40").format.font.color = palette.blue;
companies.getRange("F6:F40").format.wrapText = true;
companies.getRange("H6:H40").format.wrapText = true;
companies.freezePanes.freezeRows(5);
addListValidation(companies, "D6:D40", priorityList);

// Contacts
styleTitle(contacts, contacts.getRange("A1:H1"), "Networking & Referral Tracker", "Track warm contacts, referral asks, and follow-ups.");
contacts.getRange("A5:H5").values = [contactHeaders];
styleHeader(contacts.getRange("A5:H5"));
writeMatrix(contacts, 5, 0, contactRows);
styleTable(contacts.getRange("A6:H80"));
setWidths(contacts, [18, 24, 24, 14, 14, 14, 12, 42]);
contacts.getRange("E6:F80").setNumberFormat("yyyy-mm-dd");
contacts.getRange("H6:H80").format.wrapText = true;
contacts.freezePanes.freezeRows(5);
addListValidation(contacts, "G6:G80", ["New", "Warm", "Asked", "Referred", "Closed"]);

// Dashboard
dashboard.getRange("A1:N1").merge();
dashboard.getRange("A1:N1").values = [["Thai Big Tech Internship Dashboard"]];
dashboard.getRange("A1:N1").format.fill.color = palette.navy;
dashboard.getRange("A1:N1").format.font.color = "#FFFFFF";
dashboard.getRange("A1:N1").format.font.bold = true;
dashboard.getRange("A1:N1").format.font.size = 20;
dashboard.getRange("A1:N1").format.rowHeight = 36;
dashboard.getRange("A2:N2").merge();
dashboard.getRange("A2:N2").values = [["Snapshot, pipeline, follow-ups, and target-company coverage. Edit the Applications sheet to refresh this view."]];
dashboard.getRange("A2:N2").format.fill.color = palette.navy;
dashboard.getRange("A2:N2").format.font.color = "#DDE7F0";
dashboard.getRange("A2:N2").format.font.size = 10;
setWidths(dashboard, [15, 15, 15, 15, 3, 15, 15, 15, 15, 3, 19, 16, 14, 24]);

const kpiLabels = ["Total Roles", "Applied", "Interviews", "Offers", "Due Next 14d", "Avg Fit"];
dashboard.getRange("A4:B4").merge(); dashboard.getRange("A4:B4").values = [[kpiLabels[0]]];
dashboard.getRange("C4:D4").merge(); dashboard.getRange("C4:D4").values = [[kpiLabels[1]]];
dashboard.getRange("F4:G4").merge(); dashboard.getRange("F4:G4").values = [[kpiLabels[2]]];
dashboard.getRange("H4:I4").merge(); dashboard.getRange("H4:I4").values = [[kpiLabels[3]]];
dashboard.getRange("K4:L4").merge(); dashboard.getRange("K4:L4").values = [[kpiLabels[4]]];
dashboard.getRange("M4:N4").merge(); dashboard.getRange("M4:N4").values = [[kpiLabels[5]]];
dashboard.getRange("A5:B6").merge(); dashboard.getRange("C5:D6").merge(); dashboard.getRange("F5:G6").merge();
dashboard.getRange("H5:I6").merge(); dashboard.getRange("K5:L6").merge(); dashboard.getRange("M5:N6").merge();
dashboard.getRange("A5").formulas = [["=COUNTA('Applications'!$A$6:$A$105)"]];
dashboard.getRange("C5").formulas = [["=COUNTIF('Applications'!$F$6:$F$105,\"Applied\")"]];
dashboard.getRange("F5").formulas = [["=COUNTIF('Applications'!$F$6:$F$105,\"Interview\")"]];
dashboard.getRange("H5").formulas = [["=COUNTIF('Applications'!$F$6:$F$105,\"Offer\")"]];
dashboard.getRange("K5").formulas = [["=COUNTIFS('Applications'!$H$6:$H$105,\">=\"&TODAY(),'Applications'!$H$6:$H$105,\"<=\"&TODAY()+14,'Applications'!$F$6:$F$105,\"<>Rejected\")"]];
dashboard.getRange("M5").formulas = [["=IFERROR(AVERAGE('Applications'!$K$6:$K$105),0)"]];
dashboard.getRange("A4:N6").format.borders = { preset: "outside", style: "thin", color: palette.line };
for (const address of ["A4:B6", "C4:D6", "F4:G6", "H4:I6", "K4:L6", "M4:N6"]) {
  const block = dashboard.getRange(address);
  block.format.fill.color = palette.panel;
  block.format.borders = { preset: "outside", style: "thin", color: palette.line };
}
dashboard.getRange("A4:N4").format.font.bold = true;
dashboard.getRange("A4:N4").format.font.color = palette.muted;
dashboard.getRange("A5:N6").format.font.size = 24;
dashboard.getRange("A5:N6").format.font.bold = true;
dashboard.getRange("A5:N6").setNumberFormat("0");
dashboard.getRange("M5:N6").setNumberFormat("0.0");

dashboard.getRange("A9:C9").values = [["Pipeline Stage", "Count", "Share"]];
styleHeader(dashboard.getRange("A9:C9"));
dashboard.getRange("A10:A17").values = statusList.map((s) => [s]);
dashboard.getRange("B10:B17").formulas = statusList.map((s) => [`=COUNTIF('Applications'!$F$6:$F$105,A${10 + statusList.indexOf(s)})`]);
dashboard.getRange("C10:C17").formulas = statusList.map((_, i) => [`=IFERROR(B${10 + i}/$A$5,0)`]);
dashboard.getRange("C10:C17").setNumberFormat("0%");
styleTable(dashboard.getRange("A10:C17"));

dashboard.getRange("F9:H9").values = [["Priority", "Open Roles", "Avg Fit"]];
styleHeader(dashboard.getRange("F9:H9"));
dashboard.getRange("F10:F12").values = priorityList.map((p) => [p]);
dashboard.getRange("G10:G12").formulas = priorityList.map((_, i) => [`=COUNTIFS('Applications'!$G$6:$G$105,F${10 + i},'Applications'!$F$6:$F$105,\"<>Rejected\",'Applications'!$F$6:$F$105,\"<>Withdrawn\")`]);
dashboard.getRange("H10:H12").formulas = priorityList.map((_, i) => [`=IFERROR(SUMIFS('Applications'!$K$6:$K$105,'Applications'!$G$6:$G$105,F${10 + i})/COUNTIFS('Applications'!$G$6:$G$105,F${10 + i},'Applications'!$K$6:$K$105,\">0\"),0)`]);
dashboard.getRange("H10:H12").setNumberFormat("0.0");
styleTable(dashboard.getRange("F10:H12"));

dashboard.getRange("K9:N9").values = [["Company", "Status", "Deadline", "Next Step"]];
styleHeader(dashboard.getRange("K9:N9"));
for (let r = 10; r <= 17; r++) {
  const sourceRow = r - 4;
  dashboard.getRange(`K${r}`).formulas = [[`='Applications'!A${sourceRow}`]];
  dashboard.getRange(`L${r}`).formulas = [[`='Applications'!F${sourceRow}`]];
  dashboard.getRange(`M${r}`).formulas = [[`='Applications'!H${sourceRow}`]];
  dashboard.getRange(`N${r}`).formulas = [[`='Applications'!N${sourceRow}`]];
}
dashboard.getRange("M10:M17").setNumberFormat("yyyy-mm-dd");
dashboard.getRange("N10:N17").format.wrapText = true;
styleTable(dashboard.getRange("K10:N17"));

dashboard.getRange("A20:C20").values = [["Role Type", "Count", "Avg Fit"]];
styleHeader(dashboard.getRange("A20:C20"));
dashboard.getRange("A21:A28").values = roleTypeList.map((r) => [r]);
dashboard.getRange("B21:B28").formulas = roleTypeList.map((_, i) => [`=COUNTIF('Applications'!$C$6:$C$105,A${21 + i})`]);
dashboard.getRange("C21:C28").formulas = roleTypeList.map((_, i) => [`=IFERROR(SUMIFS('Applications'!$K$6:$K$105,'Applications'!$C$6:$C$105,A${21 + i})/COUNTIFS('Applications'!$C$6:$C$105,A${21 + i},'Applications'!$K$6:$K$105,\">0\"),0)`]);
dashboard.getRange("C21:C28").setNumberFormat("0.0");
styleTable(dashboard.getRange("A21:C28"));

dashboard.getRange("F20:H20").values = [["Source", "Applications", "Interview Rate"]];
styleHeader(dashboard.getRange("F20:H20"));
dashboard.getRange("F21:F27").values = sourceList.map((s) => [s]);
dashboard.getRange("G21:G27").formulas = sourceList.map((_, i) => [`=COUNTIF('Applications'!$E$6:$E$105,F${21 + i})`]);
dashboard.getRange("H21:H27").formulas = sourceList.map((_, i) => [`=IFERROR((COUNTIFS('Applications'!$E$6:$E$105,F${21 + i},'Applications'!$F$6:$F$105,\"Interview\")+COUNTIFS('Applications'!$E$6:$E$105,F${21 + i},'Applications'!$F$6:$F$105,\"Offer\"))/G${21 + i},0)`]);
dashboard.getRange("H21:H27").setNumberFormat("0%");
styleTable(dashboard.getRange("F21:H27"));

dashboard.getRange("K20:N20").values = [["This Week Focus", "Value", "Formula", "Meaning"]];
styleHeader(dashboard.getRange("K20:N20"));
dashboard.getRange("K21:N25").values = [
  ["High-priority open", null, "COUNTIFS", "Roles that deserve the next application push"],
  ["Follow-ups due", null, "COUNTIFS", "Open roles with next follow-up on or before today"],
  ["Need referral", null, "COUNTIFS", "High-priority rows with no referral yet"],
  ["Saved but not applied", null, "COUNTIF", "Roles to convert into applications"],
  ["Rejected / withdrawn", null, "COUNTIF", "Closed roles to learn from"],
];
dashboard.getRange("L21").formulas = [["=COUNTIFS('Applications'!$G$6:$G$105,\"High\",'Applications'!$F$6:$F$105,\"<>Rejected\",'Applications'!$F$6:$F$105,\"<>Withdrawn\")"]];
dashboard.getRange("L22").formulas = [["=COUNTIFS('Applications'!$J$6:$J$105,\">=1\",'Applications'!$J$6:$J$105,\"<=\"&TODAY(),'Applications'!$F$6:$F$105,\"<>Rejected\",'Applications'!$F$6:$F$105,\"<>Withdrawn\")"]];
dashboard.getRange("L23").formulas = [["=COUNTIFS('Applications'!$G$6:$G$105,\"High\",'Applications'!$L$6:$L$105,\"None\")"]];
dashboard.getRange("L24").formulas = [["=COUNTIF('Applications'!$F$6:$F$105,\"Saved\")"]];
dashboard.getRange("L25").formulas = [["=COUNTIF('Applications'!$F$6:$F$105,\"Rejected\")+COUNTIF('Applications'!$F$6:$F$105,\"Withdrawn\")"]];
dashboard.getRange("N21:N25").format.wrapText = true;
styleTable(dashboard.getRange("K21:N25"));

const pipelineChart = dashboard.charts.add("ColumnClustered", dashboard.getRange("A9:B17"), "Auto");
pipelineChart.title.text = "Pipeline by Stage";
pipelineChart.setPosition(dashboard.getRange("A31:D44"));
pipelineChart.width = 430;
pipelineChart.height = 260;
pipelineChart.yAxis = { numberFormatCode: "0", majorUnit: 1 };

const priorityChart = dashboard.charts.add("ColumnClustered", dashboard.getRange("F9:G12"), "Auto");
priorityChart.title.text = "Open Roles by Priority";
priorityChart.setPosition(dashboard.getRange("F31:I44"));
priorityChart.width = 410;
priorityChart.height = 260;
priorityChart.yAxis = { numberFormatCode: "0", majorUnit: 1 };

const roleChart = dashboard.charts.add("BarClustered", dashboard.getRange("A20:B28"), "Auto");
roleChart.title.text = "Role Type Mix";
roleChart.setPosition(dashboard.getRange("K31:N44"));
roleChart.width = 430;
roleChart.height = 260;
roleChart.xAxis = { numberFormatCode: "0", majorUnit: 1 };

// README
readme.getRange("A1:F1").merge();
readme.getRange("A1:F1").values = [["How to use this internship dashboard"]];
readme.getRange("A1:F1").format.fill.color = palette.navy;
readme.getRange("A1:F1").format.font.color = "#FFFFFF";
readme.getRange("A1:F1").format.font.bold = true;
readme.getRange("A1:F1").format.font.size = 18;
readme.getRange("A3:F9").values = [
  ["Step", "What to do", "Where"],
  ["1", "Add or replace target companies and roles.", "Applications"],
  ["2", "Keep status, priority, deadline, and next follow-up current.", "Applications"],
  ["3", "Use Company Targets as a shortlist and research queue.", "Company Targets"],
  ["4", "Track contacts and referral asks separately from applications.", "Contacts"],
  ["5", "Use Dashboard every week to decide what to apply to, follow up on, or close.", "Dashboard"],
  ["Note", "Sample rows are starter examples only. Replace them with live openings you find.", "All sheets"],
];
styleHeader(readme.getRange("A3:C3"));
styleTable(readme.getRange("A4:C9"));
setWidths(readme, [14, 62, 20, 12, 12, 12]);
readme.getRange("B4:B9").format.wrapText = true;

// Common visual polish
for (const sheet of [dashboard, apps, companies, contacts, lookups, readme]) {
  const used = sheet.getUsedRange();
  used.format.font.name = "Aptos";
}

await fs.mkdir(outputDir, { recursive: true });

const dashInspect = await workbook.inspect({
  kind: "table",
  sheetId: "Dashboard",
  range: "A1:N28",
  include: "values,formulas",
  tableMaxRows: 30,
  tableMaxCols: 14,
  maxChars: 5000,
});
console.log(dashInspect.ndjson);

const errors = await workbook.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
  options: { useRegex: true, maxResults: 300 },
  summary: "final formula error scan",
  maxChars: 4000,
});
console.log(errors.ndjson);

for (const sheetName of ["Dashboard", "Applications", "Company Targets", "Contacts", "Lookups", "README"]) {
  const preview = await workbook.render({ sheetName, autoCrop: "all", scale: 1, format: "png" });
  const bytes = new Uint8Array(await preview.arrayBuffer());
  await fs.writeFile(`${outputDir}/preview_${sheetName.replaceAll(" ", "_")}.png`, bytes);
}

const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(outputPath);
console.log(`SAVED ${outputPath}`);
