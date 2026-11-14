"use strict";
// Cấu hình nguồn. Tên file giữ cố định khi bạn thay bản Excel trong thư mục.
const sources = ["HPDQ1", "HPDQ2"];
const $ = id => document.getElementById(id);
const clean = value => String(value ?? "").trim().replace(/\s+/g, " ");
const norm = value => clean(value).toLocaleLowerCase("vi").normalize("NFD").replace(/\p{M}/gu, "").replace(/đ/g, "d");
const has = value => clean(value) !== "";
const cabinetNames = new Set(["ten tu", "ma tu", "ky hieu tu", "feeder no", "cabinet", "tu trung the"]);
const simpleTabs = new Set(["Bãi Liệu", "Luyện Cốc", "TKVV", "LT", "LG1", "CTD1", "CTD2", "CTD3", "QSP", "Sheet1"]);
const specialSections = /^(?:ccm\d|rhf \(|dc 10|xln \(|khu khi khoi$|phat dien nhiet du$|quat gio$|tinh luyen lf)/;
const regionTree = new Map();
let current = null;

function element(tag, text, parent, className) {
  const node = document.createElement(tag);
  if (text != null) node.textContent = text;
  if (className) node.className = className;
  if (parent) parent.append(node);
  return node;
}

function headerAt(row) {
  const labels = row.map(norm);
  const col = labels.findIndex(label => cabinetNames.has(label));
  if (col < 0) return null;
  const load = labels.findIndex((label,i) => i > col && /^(?:ten phu tai|ten tai|lo \(phu tai\)|phu tai|feeder|id)$/.test(label));
  return {col, load};
}

function sectionName(row, heading) {
  if (heading) return "";
  const values = row.filter(has);
  if (values.length !== 1 || typeof values[0] !== "string") return "";
  const label = clean(values[0]);
  if (/^(?:[ivx]+\s*\.|\d+\s*\.)\s*/i.test(label) || /^(?:tram(?:\s|$)|tu 35|tu 10)/.test(norm(label)) || specialSections.test(norm(label))) return label;
  return "";
}

// Trải tiêu đề qua ô gộp. Tiêu đề theo từng cột gốc để tránh ghép nhầm đơn vị.
function makeHeaders(sheet, start, end, width) {
  const rowValues = (r,c) => {
    if (has(sheet.rows[r]?.[c])) return sheet.rows[r][c];
    const merge = sheet.merges.find(([[r1,c1],[r2,c2]]) => r >= r1 && r <= r2 && c+1 >= c1 && c+1 <= c2);
    return merge ? sheet.rows[merge[0][0]]?.[merge[0][1]-1] : "";
  };
  return Array.from({length:width}, (_,col) => {
    const parts = [];
    for (let r=start; r<=end; r++) {
      const value = clean(rowValues(r,col));
      if (value && parts.at(-1) !== value) parts.push(value);
    }
    return parts;
  });
}

function parse(sheet) {
  const blocks = [];
  let active = null;
  let head = null;
  let label = sheet.name === "LG1" ? "Xuất tuyến 110 kV" : sheet.name;
  let sectionRow = 1;
  const rows = sheet.rows;
  const cutoff = sheet.name === "Bãi Liệu" ? rows.findIndex(row => row?.some(value => /^quang ngai, ngay/.test(norm(value)))) : -1;
  function setHeader(r, info, depth) {
    head = { ...info, end:r+depth-1, columns:makeHeaders(sheet,r,r+depth-1,sheet.width) };
    active = null;
  }
  function add(r, cells) {
    if (!active) {
      active = {label, sectionRow, col:head.col, load:head.load, headers:head.columns, rows:[], sheet:sheet.name};
      blocks.push(active);
    }
    active.rows.push({r,cells, dates:[...sheet.dateCells].filter(ref => Number(ref.match(/\d+$/)[0]) === r)});
  }
  // Sheet tổng hợp chạm đất có hai bảng nằm cạnh nhau. Tách theo vùng cột.
  if (sheet.name === "Sheet1" && headerAt(rows[3] || []) && headerAt((rows[3] || []).slice(10))) {
    for (const [start,end,titleRow,headerRow,last] of [[0,10,2,3,19],[10,sheet.width,2,3,rows.length-1],[0,10,20,3,rows.length-1]]) {
      const header = rows[headerRow]?.slice(start,end) ?? [];
      if (!headerAt(header)) continue;
      label = clean(rows[titleRow]?.[start]);
      sectionRow = titleRow;
      head = {col:0, load:1, columns:Array.from({length:end-start}, (_,i) => has(header[i]) ? [clean(header[i])] : [])};
      active = null;
      for (let r=Math.max(headerRow,titleRow)+1;r<=last;r++) {
        const cells = rows[r]?.slice(start,end);
        if (cells?.some(has)) add(r,cells);
      }
      for (const item of active?.rows ?? []) item.dates = [];
      if (active) active.offset = start;
    }
    return blocks;
  }

  for (let r=1;r<(cutoff > 0 ? cutoff : rows.length);r++) {
    const cells = rows[r];
    if (!cells?.some(has)) continue;
    if (head && r <= head.end) continue;
    const heading = headerAt(cells);
    const section = sectionName(cells, heading);
    if (section) {
      label = section;
      sectionRow = r;
      active = null;
      continue;
    }
    if (heading) {
      let depth = 2;
      if (simpleTabs.has(sheet.name)) {
        depth = sheet.name === "Sheet1" ? 1 : 3;
        if (sheet.name === "TKVV" && heading.col === 1) {
          label = "Bảng bổ sung";
          sectionRow = r;
          depth = 1;
        }
      }
      setHeader(r, heading, depth);
      continue;
    }
    if (!head) continue;
    if (cells.filter(has).length === 1 && !has(cells[head.col])) continue;
    if (cells.some(value => /^(?:quang ngai, ngay|ngay \d+ thang|p\.tb$)/.test(norm(value)))) continue;
    // Hai nhóm cuối LT2 đổi từ bố cục STT sang bố cục tủ ở cột A, thiếu tiêu đề riêng.
    if (sheet.name === "LT2" && /svg#|tinh luyen lf/.test(norm(label)) && head.col !== 0) {
      head = {col:0,load:1,end:r-1,columns:Array.from({length:sheet.width}, (_,i) => i === 0 ? ["Tên / mã tủ"] : [])};
    }
    if (sheet.name === "LG1" && /^110\s*kv$/i.test(clean(cells[0]))) {
      if (!active) head = {...head,col:1,load:-1};
    } else if (sheet.name === "LG1" && head.col === 1 && !active) {
      head = {...head,col:0,load:1};
    }
    add(r,cells);
  }
  return blocks;
}

function arrange(blocks) {
  const clusters = new Map();
  let parent = "";
  for (const block of blocks) {
    const name = norm(block.label);
    const main = /35\s*kv/.test(name) && !/\b(?:10|3)\s*kv|sub\d|er\d/.test(name);
    const transmission = /110\s*kv/.test(name);
    if (main) parent = block.label;
    const key = transmission ? "Trạm 110 kV" : parent || "Chưa xác định trạm 35 kV";
    if (!clusters.has(key)) clusters.set(key, []);
    const items = clusters.get(key);
    block.display = main ? "Bảng chỉnh định · " + block.label : block.label;
    if (items.some(item => item.display === block.display)) block.display += " · dòng " + block.sectionRow;
    items.push(block);
  }
  return clusters;
}

function branch(parent, label, className) {
  const node = element("details", null, parent, className);
  element("summary", label, node);
  return element("div", null, node, "branch");
}

function columnsFor(block) {
  const indices = new Set([block.col]);
  if (block.load >= 0) indices.add(block.load);
  block.headers.forEach((label,col) => {
    if (label.length || block.rows.some(row => has(row.cells[col]))) indices.add(col);
  });
  return [...indices];
}

function render() {
  if (!current) return;
  const {block} = current;
  const term = norm($("search").value);
  const rows = block.rows.filter(row => norm(row.cells[block.col]).includes(term));
  const table = $("table");
  table.replaceChildren();
  const cols = columnsFor(block);
  const header = table.createTHead().insertRow();
  for (const col of cols) {
    const labels = block.headers[col] || [];
    const heading = col === block.col ? "Tên / mã tủ" : col === block.load ? "Phụ tải" : labels.at(-1) || "Cột " + Excel.colName(col + 1 + (block.offset || 0));
    const cell = element("th", heading, header);
    cell.scope = "col";
    const context = col === block.col || col === block.load ? labels : labels.slice(0,-1);
    if (context.length) element("small", context.join(" · "), cell);
  }
  const body = table.createTBody();
  for (const row of rows) {
    const tr = body.insertRow();
    tr.title = "Dòng " + row.r + " trong Excel";
    for (const col of cols) {
      const date = row.dates.includes(Excel.colName(col+1+(block.offset || 0)) + row.r);
      const value = row.cells[col];
const displayValue =
  typeof value === "number" && col !== block.col
    ? value.toFixed(2).replace(/\.?0+$/, "")
    : (value ?? "");

const td = element("td", displayValue, tr, date ? "date" : undefined);
      if (date) td.title = "Ô có định dạng ngày trong file nguồn. Đối chiếu nếu đây là thông số chỉnh định.";
    }
  }
  $("count").textContent = rows.length + " / " + block.rows.length + " dòng";
  $("empty").hidden = rows.length > 0;
}

function select(entry, button) {
  document.querySelector(".leaf[aria-current]")?.removeAttribute("aria-current");
  button.setAttribute("aria-current", "true");
  for (let node=button.parentElement; node; node=node.parentElement) if (node.tagName === "DETAILS") node.open = true;
  current = entry;
  $("path").textContent = [entry.region, entry.unit, entry.station].join(" / ");
  $("title").textContent = entry.block.display;
  $("source").textContent = "Nguồn " + entry.region + ".xlsx · Sheet " + entry.block.sheet + " · Từ dòng " + entry.block.rows[0].r;
  $("viewer").hidden = false;
  $("search").value = "";
  $("export-group").textContent = entry.unit === "Chưa xác định NM.CĐ" ? "Xuất Excel sheet nguồn" : "Xuất Excel nhà máy";
  $("table").parentElement.scrollTo(0,0);
  render();
}

function addGroup(region, sheets) {
  const regionNode = branch($("tree"),region,"region");
  let firstButton;
  // Gán NM.CĐ chỉ khi file nêu rõ. Có thể điền thêm bảng ánh xạ khi xác nhận.
  const confirmedUnits = region === "HPDQ1" ? {"Luyện Cốc":"NM.CĐ1", "Bãi Liệu":"NM.CĐ1"} : {};
  const units = new Map();
  for (const sheet of sheets) {
    const blocks = parse(sheet);
    if (!blocks.length) continue;
    const unit = confirmedUnits[sheet.name] || "Chưa xác định NM.CĐ";
    const clusters = arrange(blocks);
    if (!units.has(unit)) {
      const unitNode = branch(regionNode, unit);
      unitNode.parentElement.open = true;
      units.set(unit, unitNode);
    }
    const sheetNode = branch(units.get(unit),sheet.name);
    for (const [station, list] of clusters) {
      const stationNode = branch(sheetNode,station);
      for (const block of list) {
        const button = element("button",block.display,stationNode,"leaf");
        button.addEventListener("click", () => select({block,region,unit,station,group:unit === "Chưa xác định NM.CĐ" ? blocks : regionTree.get(region).get(unit)},button));
        firstButton ??= button;
      }
    }
    const groups = regionTree.get(region);
    if (!groups.has(unit)) groups.set(unit,[]);
    groups.get(unit).push(...blocks);
  }
  return firstButton;
}

function notice(message) {
  $("notice").textContent = message;
  $("notice").hidden = !message;
}

async function download(blocks, filename) {
  try {
    const blob = await Excel.write(blocks.map(block => {
      const cols = columnsFor(block);
      return {name:block.label, rows:[
        cols.map(col => block.headers[col]?.join(" / ") || (col === block.col ? "Tên / mã tủ" : "Cột " + Excel.colName(col+1+(block.offset||0)))),
        ...block.rows.map(row => cols.map(col => row.cells[col]))
      ]};
    }));
    const url = URL.createObjectURL(blob);
    const link = element("a",null,document.body);
    link.href = url;
    link.download = filename.replace(/[\\/:*?"<>|]/g,"-") + ".xlsx";
    link.click();
    link.remove();
    setTimeout(() => URL.revokeObjectURL(url),3000);
  } catch (error) {
    notice("Không xuất được Excel. " + error.message);
  }
}

$("search").addEventListener("input",render);
$("export-station").addEventListener("click",() => download([current.block],current.region+" - "+current.block.label));
$("export-group").addEventListener("click",() => download(current.group,current.region+" - "+(current.unit === "Chưa xác định NM.CĐ" ? current.block.sheet : current.unit)));
function setTheme(dark) {
  document.documentElement.classList.toggle("dark",dark);
  $("theme").textContent = dark ? "☀ Chế độ sáng" : "◐ Chế độ tối";
  $("theme").setAttribute("aria-pressed",String(dark));
}
try { setTheme(localStorage.getItem("hpdq-theme") === "dark"); } catch {}
$("theme").addEventListener("click",() => {
  const dark = !document.documentElement.classList.contains("dark");
  setTheme(dark);
  try { localStorage.setItem("hpdq-theme",dark ? "dark" : "light"); } catch {}
});

// Bản xem thử mang sẵn hai tệp. Bản mã nguồn đọc tệp thật trong du-lieu khi mở web.
async function loadSource(region) {
  if (window.RELAY_PREVIEW?.[region]) return Excel.read(Uint8Array.from(atob(window.RELAY_PREVIEW[region]), ch => ch.charCodeAt(0)));
  const response = await fetch("du-lieu/" + region + ".xlsx",{cache:"no-store"});
  if (!response.ok) throw new Error("HTTP " + response.status);
  return Excel.read(await response.arrayBuffer());
}
async function start() {
  notice("Đang đọc hai file Excel…");
  const results = await Promise.allSettled(sources.map(loadSource));
  let first;
  const errors = [];
  results.forEach((result,i) => {
    const region = sources[i];
    if (result.status === "rejected") {
      errors.push(region);
      console.error(region,result.reason);
      return;
    }
    regionTree.set(region,new Map());
    const firstInRegion = addGroup(region,result.value);
    first ??= firstInRegion;
  });
  if (first) first.click();
  notice(errors.length ? "Chưa đọc được " + errors.join(", ") + ". Kiểm tra thư mục du-lieu và mở bản mã nguồn qua máy chủ web." : "");
}
start().catch(error => { console.error(error); notice("Không thể mở dữ liệu. " + error.message); });



const value = row.cells[col];
const displayValue =
  typeof value === "number" && col !== block.col
    ? value.toFixed(2).replace(/\.00$/, "")
    : (value ?? "");