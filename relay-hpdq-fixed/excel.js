"use strict";

// Đọc định dạng XLSX: mỗi bảng tính là các tài liệu XML trong một gói ZIP.
const Excel = (() => {
  const relation = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
  const spreadsheet = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
  const get = (node, tag) => Array.from(node.getElementsByTagNameNS("*", tag));
  const cellText = node => get(node, "t").filter(t => t.parentNode.localName !== "rPh").map(t => t.textContent).join("");
  function coordinate(ref) {
    const match = /([A-Z]+)(\d+)/.exec(ref);
    let col = 0;
    for (const char of match[1]) col = col * 26 + char.charCodeAt(0) - 64;
    return [Number(match[2]), col];
  }
  function colName(col) {
    let name = "";
    while (col) {
      col--;
      name = String.fromCharCode(65 + col % 26) + name;
      col = Math.floor(col / 26);
    }
    return name;
  }
  function dateText(value, format, base1904) {
    if (!Number.isFinite(value)) return String(value);
    const date = new Date(Date.UTC(base1904 ? 1904 : 1899, base1904 ? 0 : 11, base1904 ? 1 : 30) + value * 86400000);
    const iso = date.toISOString();
    if (/^m[\/-]yy$/i.test(format)) return (date.getUTCMonth() + 1) + format[1] + iso.slice(2, 4);
    return iso.split("T")[0];
  }
  async function read(bytes) {
    const zip = await JSZip.loadAsync(bytes);
    async function xml(path) {
      const file = zip.file(path);
      if (!file) return null;
      const doc = new DOMParser().parseFromString(await file.async("string"), "application/xml");
      if (get(doc, "parsererror").length) throw new Error("Không đọc được " + path);
      return doc;
    }
    const book = await xml("xl/workbook.xml");
    const links = await xml("xl/_rels/workbook.xml.rels");
    const targets = new Map(get(links, "Relationship").map(x => [x.getAttribute("Id"), new URL(x.getAttribute("Target"), "https://xlsx.invalid/xl/workbook.xml").pathname.slice(1)]));
    const shared = await xml("xl/sharedStrings.xml");
    const strings = shared ? get(shared, "si").map(cellText) : [];
    const styles = await xml("xl/styles.xml");
    const custom = new Map(styles ? get(styles, "numFmt").map(n => [Number(n.getAttribute("numFmtId")), n.getAttribute("formatCode")]) : []);
    const xfs = styles ? get(styles, "cellXfs")[0] : null;
    const formats = xfs ? Array.from(xfs.children).map(n => Number(n.getAttribute("numFmtId"))) : [];
    const base1904 = ["1", "true"].includes(get(book, "workbookPr")[0]?.getAttribute("date1904"));
    return Promise.all(get(book, "sheet").map(async item => {
      const sheet = await xml(targets.get(item.getAttributeNS(relation, "id")));
      const rows = [];
      const dateCells = new Set();
      let width = 0;
      for (const row of get(sheet, "row")) {
        const number = Number(row.getAttribute("r"));
        const values = [];
        for (const cell of Array.from(row.children).filter(n => n.localName === "c")) {
          const ref = cell.getAttribute("r");
          const [, col] = coordinate(ref);
          const type = cell.getAttribute("t");
          const raw = get(cell, "v")[0]?.textContent ?? "";
          let value = raw;
          if (type === "inlineStr") value = cellText(cell);
          else if (type === "s") value = strings[Number(raw)] ?? "";
          else if (type === "b") value = raw === "1";
          else if ((!type || type === "n") && raw !== "") {
            value = Number(raw);
            const id = formats[Number(cell.getAttribute("s"))];
            const format = custom.get(id) ?? "";
            if ((id >= 14 && id <= 22) || /[dy]/i.test(format.replace(/"[^"]*"|\\./g, ""))) {
              value = dateText(value, format, base1904);
              dateCells.add(ref);
            }
          }
          if (value === "" && get(cell, "f").length) value = "[Công thức chưa có kết quả lưu]";
          if (value !== "") {
            values[col - 1] = value;
            width = Math.max(width, col);
          }
        }
        if (values.length) rows[number] = values;
      }
      const merges = get(sheet, "mergeCell").map(node => node.getAttribute("ref").split(":").map(coordinate));
      return {name:item.getAttribute("name"), rows, width, merges, dateCells};
    }));
  }

  // Tạo tệp XLSX thật. Văn bản được ghi dưới dạng chuỗi, số giữ kiểu số.
  const escape = value => String(value).replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f]/g, "").replace(/[&<>"']/g, c => ({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&apos;"})[c]);
  async function write(tabs) {
    const zip = new JSZip();
    const declarations = '<?xml version="1.0" encoding="UTF-8"?>';
    const relns = "http://schemas.openxmlformats.org/package/2006/relationships";
    const contentTypes = "application/vnd.openxmlformats-officedocument.spreadsheetml";
    const names = new Set();
    const sheets = tabs.map((tab, index) => {
      const base = tab.name.replace(/[\\/?*:[\]]/g, " ").replace(/^'+|'+$/g, "").trim().slice(0, 31) || "Trạm";
      let name = base;
      let suffix = 2;
      while (names.has(name.toLowerCase())) {
        const ending = " (" + suffix++ + ")";
        name = base.slice(0, 31 - ending.length) + ending;
      }
      names.add(name.toLowerCase());
      const rows = tab.rows.map((row, rowIndex) => {
        const number = rowIndex + 1;
        const cells = row.map((value, col) => {
          if (value === "" || value == null) return "";
          const ref = colName(col + 1) + number;
          if (typeof value === "number" && Number.isFinite(value)) return '<c r="' + ref + '"><v>' + value + '</v></c>';
          if (typeof value === "boolean") return '<c r="' + ref + '" t="b"><v>' + Number(value) + '</v></c>';
          return '<c r="' + ref + '" t="inlineStr"><is><t xml:space="preserve">' + escape(value) + '</t></is></c>';
        }).join("");
        return '<row r="' + number + '">' + cells + '</row>';
      }).join("");
      zip.file("xl/worksheets/sheet" + (index + 1) + ".xml", declarations + '<worksheet xmlns="' + spreadsheet + '"><sheetData>' + rows + '</sheetData></worksheet>');
      return '<sheet name="' + escape(name) + '" sheetId="' + (index + 1) + '" r:id="rId' + (index + 1) + '"/>';
    });
    zip.file("[Content_Types].xml", declarations + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="' + contentTypes + '.sheet.main+xml"/>' + tabs.map((_,i) => '<Override PartName="/xl/worksheets/sheet' + (i+1) + '.xml" ContentType="' + contentTypes + '.worksheet+xml"/>').join("") + '</Types>');
    zip.file("_rels/.rels", declarations + '<Relationships xmlns="' + relns + '"><Relationship Id="rId1" Type="' + relation + '/officeDocument" Target="xl/workbook.xml"/></Relationships>');
    zip.file("xl/workbook.xml", declarations + '<workbook xmlns="' + spreadsheet + '" xmlns:r="' + relation + '"><sheets>' + sheets.join("") + '</sheets></workbook>');
    zip.file("xl/_rels/workbook.xml.rels", declarations + '<Relationships xmlns="' + relns + '">' + tabs.map((_,i) => '<Relationship Id="rId' + (i+1) + '" Type="' + relation + '/worksheet" Target="worksheets/sheet' + (i+1) + '.xml"/>').join("") + '</Relationships>');
    return zip.generateAsync({type:"blob", compression:"DEFLATE", mimeType:"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"});
  }
  return {read, write, colName};
})();
