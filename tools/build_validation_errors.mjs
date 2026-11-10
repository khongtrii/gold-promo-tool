import fs from "node:fs/promises";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const outputDir = "D:/Code/gold-promo-tool/outputs/validation-errors";
const outputPath = `${outputDir}/gold_promo_validation_errors.xlsx`;
const previewPath = `${outputDir}/gold_promo_validation_errors.png`;

const rows = [
  ["Các cột bắt buộc đang để trống: {danh sách cột}", "ATTRIBUTE MARKETING; GOLD PROMO TYPE; GOLD CODE; LV; LU; SUPPLIER CODE; COMMERCIAL CONTRACT; PURCHASE NETWORK; GOLD PROMO NETWORK; NORMAL PURCHASE PRICE; PURCHASE VAT; DISCOUNT (% OR VALUE); DELIVERY TYPE; PP START DAY/MONTH/YEAR; PP END DAY/MONTH/YEAR; % DELIVERY 1/2/3. ATTRIBUTE MARKETING chỉ bắt buộc khi bật Attribute; FREE PRODUCT được kiểm tra bằng rule riêng.", "Stage 1; Stage 1 - Attribute"],
  ["Các cột bắt buộc đang để trống: {danh sách cột}", "GOLD PROMO TYPE; GOLD CODE; SV; PRICELIST; SALE VAT; SP START DAY/MONTH/YEAR; SP END DAY/MONTH/YEAR", "Stage 2"],
  ["Các cột bắt buộc đang để trống: {danh sách cột}", "GOLD PROMO TYPE; GOLD CODE; LV; SUPPLIER CODE; COMMERCIAL CONTRACT; PURCHASE NETWORK; DISCOUNT (% OR VALUE); PP START DAY/MONTH/YEAR; PP END DAY/MONTH/YEAR", "Stage 3"],
  ["Các cột bắt buộc đang để trống: {danh sách cột}", "GOLD CODE; SV; REGION (NORTH/SOUTH/CENTER/ALL); PAGE; THEMATIC; POSITION", "Stage 1 - Attribute"],
  ["GOLD PROMO TYPE chỉ được phép là GOLD PROMO hoặc OUT OF GOLD PROMO.", "GOLD PROMO TYPE", "Stage 1; Stage 2; Stage 3"],
  ["FREE PRODUCT không được để trống khi DISCOUNT (% OR VALUE) có dữ liệu.", "FREE PRODUCT; DISCOUNT (% OR VALUE)", "Stage 1 - Attribute"],
  ["Vui lòng bổ sung prefix cho ATTRIBUTE đặc biệt", "ATTRIBUTE MARKETING", "Stage 1 - Attribute"],
  ["POSITION không thể chuyển đổi thành CLASS", "POSITION; CLASS (dẫn xuất)", "Stage 1 - Attribute"],
  ["PURCHASE VAT chỉ được phép là 0%, 5%, 8%, 10%, KKKT hoặc KCT.", "PURCHASE VAT", "Stage 1"],
  ["NORMAL PURCHASE PRICE không thể chuyển đổi thành số", "NORMAL PURCHASE PRICE", "Stage 1"],
  ["DISCOUNT (% OR VALUE) không thể chuyển đổi thành số", "DISCOUNT (% OR VALUE)", "Stage 1; Stage 3"],
  ["DISCOUNT (% OR VALUE) không được vượt quá 100%.", "DISCOUNT (% OR VALUE)", "Stage 1; Stage 3"],
  ["GOLD CODE ngoại lệ không được cài chiết khấu thùng (T/TH).", "GOLD CODE; DISCOUNT (% OR VALUE); danh sách GOLD CODE ngoại lệ", "Stage 1"],
  ["Stage 3 không cho phép DISCOUNT dạng số+số; chỉ chấp nhận số, số% hoặc sốT/TH+sốT/TH.", "DISCOUNT (% OR VALUE)", "Stage 3"],
  ["Không được sử dụng network 8000.", "PURCHASE NETWORK; GOLD PROMO NETWORK; FILE NAME; GOLD CODE; LV; LU", "Stage 1"],
  ["PURCHASE NETWORK không đúng SITE RULE.", "PURCHASE NETWORK", "Stage 1"],
  ["GOLD PROMO NETWORK không đúng SITE RULE.", "GOLD PROMO NETWORK", "Stage 1"],
  ["{PURCHASE NETWORK hoặc GOLD PROMO NETWORK} sau khi expand +/- bị rỗng, vui lòng kiểm tra lại.", "PURCHASE NETWORK hoặc GOLD PROMO NETWORK; trường NETWORK EXPANDED tương ứng", "Stage 1"],
  ["{PURCHASE NETWORK hoặc GOLD PROMO NETWORK} có site trừ không thuộc network tổng: {danh sách site}.", "PURCHASE NETWORK hoặc GOLD PROMO NETWORK; network-configure; trường NETWORK EXPANDED tương ứng", "Stage 1"],
  ["Cửa hàng {danh sách site} không thuộc network 8200 hiện tại", "PURCHASE NETWORK EXPANDED; GOLD PROMO NETWORK EXPANDED; network 8200", "Stage 1"],
  ["Vui lòng kiểm tra lại GOLD PROMO NETWORK", "FILE NAME; GOLD CODE; LV; LU; GOLD PROMO NETWORK EXPANDED", "Stage 1"],
  ["PURCHASE NETWORK bị trùng lặp: {danh sách site}", "GOLD CODE; LV; PURCHASE NETWORK; PURCHASE NETWORK EXPANDED", "Stage 1"],
  ["Thiếu cửa hàng trong PURCHASE NETWORK: {danh sách site}", "FILE NAME; GOLD CODE; LV; LU; PURCHASE NETWORK EXPANDED; GOLD PROMO NETWORK EXPANDED", "Stage 1"],
  ["Dư cửa hàng trong PURCHASE NETWORK: {danh sách site}", "FILE NAME; GOLD CODE; LV; LU; PURCHASE NETWORK EXPANDED; GOLD PROMO NETWORK EXPANDED", "Stage 1"],
  ["SO chỉ được tối đa 10 ký tự.", "SO (dẫn xuất)", "Stage 1"],
  ["LU chỉ được phép là 1 hoặc 41.", "LU", "Stage 1"],
  ["DELIVERY TYPE chỉ được phép là CROSS-DOCKING, DIRECT hoặc VINAMILK.", "DELIVERY TYPE", "Stage 1"],
  ["GOLD CODE và LV trùng nhau nhưng STRUCTURE khác nhau.", "GOLD CODE; LV; STRUCTURE", "Stage 1; Stage 2"],
  ["% DELIVERY 1, % DELIVERY 2 và % DELIVERY 3 phải là số.", "% DELIVERY 1; % DELIVERY 2; % DELIVERY 3", "Stage 1"],
  ["Tổng % DELIVERY 1 + % DELIVERY 2 + % DELIVERY 3 phải bằng 100.", "% DELIVERY 1; % DELIVERY 2; % DELIVERY 3", "Stage 1"],
  ["Thông tin phân bổ phải là số nguyên và lớn hơn 0: {danh sách cửa hàng}.", "GOLD CODE; LV; LU; PURCHASE NETWORK EXPANDED; các cột phân bổ theo site", "Stage 1"],
  ["Thiếu phân bổ đối với các cửa hàng: {danh sách cửa hàng} dựa trên PURCHASE NETWORK EXPANDED.", "GOLD CODE; LV; LU; PURCHASE NETWORK EXPANDED; các cột phân bổ theo site", "Stage 1"],
  ["PP START DATE không hợp lệ.", "PP START DAY; PP START MONTH; PP START YEAR; PP START DATE (dẫn xuất)", "Stage 1"],
  ["PP END DATE không hợp lệ.", "PP END DAY; PP END MONTH; PP END YEAR; PP END DATE (dẫn xuất)", "Stage 1"],
  ["PP START DATE phải nhỏ hơn hoặc bằng PP END DATE.", "PP START DATE; PP END DATE", "Stage 1"],
  ["PP START DATE phải bằng {DELIVERY TYPE} | ORDER DATE 1 ({ngày}).", "DELIVERY TYPE; PP START DATE; {DELIVERY TYPE} | ORDER DATE 1 trong Master data", "Stage 1"],
  ["PP END DATE phải lớn hơn hoặc bằng {DELIVERY TYPE} | ORDER DATE 3 ({ngày}).", "DELIVERY TYPE; PP END DATE; {DELIVERY TYPE} | ORDER DATE 3 trong Master data", "Stage 1"],
  ["Thông tin mua hàng và chiết khấu bị trùng.", "GOLD CODE; LV; LU; SUPPLIER CODE; COMMERCIAL CONTRACT; PURCHASE NETWORK; NORMAL PURCHASE PRICE; DISCOUNT (% OR VALUE)", "Stage 1"],
  ["Thông tin mua hàng và chiết khấu bị trùng.", "GOLD CODE; LV; PURCHASE NETWORK EXPANDED; NORMAL PURCHASE PRICE; DISCOUNT (% OR VALUE)", "Stage 1"],
  ["SALE VAT chỉ được phép là 0%, 5%, 8%, 10%, KKKT hoặc KCT.", "SALE VAT", "Stage 2"],
  ["PROMOTION SALE PRICE phải là số nguyên.", "PROMOTION SALE PRICE", "Stage 2"],
  ["PRICELIST phải bắt đầu bằng 1090, 2010, 2030 hoặc 2050.", "PRICELIST", "Stage 2"],
  ["NORMAL SALE PRICE không được để trống khi PRICELIST bắt đầu bằng 1090.", "PRICELIST; NORMAL SALE PRICE", "Stage 2"],
  ["PROMOTION SALE PRICE không được để trống khi PRICELIST bắt đầu bằng 2010, 2030 hoặc 2050.", "PRICELIST; PROMOTION SALE PRICE", "Stage 2"],
  ["Cửa hàng {danh sách site} không tồn tại.", "PRICELIST; network-configure/store", "Stage 2"],
  ["SP START DATE không hợp lệ.", "SP START DAY; SP START MONTH; SP START YEAR; SP START DATE (dẫn xuất)", "Stage 2"],
  ["SP END DATE không hợp lệ.", "SP END DAY; SP END MONTH; SP END YEAR; SP END DATE (dẫn xuất)", "Stage 2"],
  ["SP START DATE phải nhỏ hơn hoặc bằng SP END DATE.", "SP START DATE; SP END DATE", "Stage 2"],
  ["SP START DATE phải bằng CATALOGUE START DATE ({ngày}).", "PRICELIST CODE; SP START DATE; CATALOGUE START DATE trong Master data", "Stage 2"],
  ["SP END DATE phải bằng CATALOGUE END DATE ({ngày}).", "PRICELIST CODE; SP END DATE; CATALOGUE END DATE trong Master data", "Stage 2"],
  ["PURCHASE NETWORK có site trừ không thuộc network tổng: {danh sách site}.", "PURCHASE NETWORK; DISCOUNT_NETWORK; DISCOUNT_NETWORK_EXPANDED", "Stage 3"],
  ["PURCHASE NETWORK sau khi expand theo DISCOUNT_NETWORK bị rỗng.", "PURCHASE NETWORK; DISCOUNT_NETWORK_EXPANDED", "Stage 3"],
  ["Site {danh sách site} không thuộc WH (8300) của DISCOUNT_NETWORK.", "PURCHASE NETWORK; DISCOUNT_NETWORK_EXPANDED; DISCOUNT_NETWORK/8300", "Stage 3"],
];

const workbook = Workbook.create();
const sheet = workbook.worksheets.add("Validation Errors");
sheet.showGridLines = false;
sheet.tabColor = "#1F4E78";

sheet.getRange("A1:C1").values = [["Danh sách lỗi validation Gold Promo", null, null]];
const headerRow = 3;
const dataStart = 4;
const dataEnd = dataStart + rows.length - 1;
sheet.getRange(`A${headerRow}:C${headerRow}`).values = [[
  "Lỗi trả về",
  "Trường thông tin được check",
  "Stage sử dụng",
]];
sheet.getRange(`A${dataStart}`).write(rows);

const used = sheet.getRange(`A1:C${dataEnd}`);
used.format.font = { name: "Arial", size: 10, color: "#222222" };
used.format.verticalAlignment = "top";
sheet.getRange("A1:C1").format.font = { name: "Arial", size: 14, bold: true, color: "#1F1F1F" };
sheet.getRange("A1:C1").format.rowHeight = 25;

const header = sheet.getRange(`A${headerRow}:C${headerRow}`);
header.format = {
  fill: "#1F4E78",
  font: { name: "Arial", size: 10, bold: true, color: "#FFFFFF" },
  horizontalAlignment: "center",
  verticalAlignment: "center",
  wrapText: true,
  borders: { preset: "all", style: "thin", color: "#FFFFFF" },
  rowHeight: 30,
};

const body = sheet.getRange(`A${dataStart}:C${dataEnd}`);
body.format.wrapText = true;
body.format.verticalAlignment = "top";
body.format.borders = {
  insideHorizontal: { style: "thin", color: "#D9E2F3" },
  bottom: { style: "thin", color: "#9EADBD" },
};

for (let row = dataStart; row <= dataEnd; row += 1) {
  if ((row - dataStart) % 2 === 1) {
    sheet.getRange(`A${row}:C${row}`).format.fill = "#F7F9FC";
  }
}

sheet.getRange(`C${dataStart}:C${dataEnd}`).format = {
  fill: "#EAF2F8",
  font: { name: "Arial", size: 10, bold: true, color: "#1F4E78" },
  horizontalAlignment: "center",
  verticalAlignment: "top",
  wrapText: true,
};

sheet.getRange("A:A").format.columnWidth = 58;
sheet.getRange("B:B").format.columnWidth = 78;
sheet.getRange("C:C").format.columnWidth = 24;
body.format.autofitRows();
sheet.freezePanes.freezeRows(headerRow);

const table = sheet.tables.add(`A${headerRow}:C${dataEnd}`, true, "ValidationErrorsTable");
table.style = "TableStyleMedium2";
table.showFilterButton = true;
table.showBandedRows = false;

workbook.recalculate();

const inspection = await workbook.inspect({
  kind: "table",
  range: `Validation Errors!A1:C${dataEnd}`,
  include: "values,formulas",
  tableMaxRows: rows.length + 4,
  tableMaxCols: 3,
  maxChars: 30000,
});
console.log(inspection.ndjson);

const errors = await workbook.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!|#NULL!|#SPILL!|#CALC!",
  options: { useRegex: true, maxResults: 100 },
  summary: "final formula error scan",
});
console.log(errors.ndjson);

await fs.mkdir(outputDir, { recursive: true });
const preview = await workbook.render({
  sheetName: "Validation Errors",
  range: `A1:C${Math.min(dataEnd, 18)}`,
  scale: 1,
  format: "png",
});
await fs.writeFile(previewPath, new Uint8Array(await preview.arrayBuffer()));

const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(outputPath);
console.log(JSON.stringify({ outputPath, previewPath, rowCount: rows.length }));
