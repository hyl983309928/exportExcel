/* eslint-disable */

import FileSaver from 'file-saver'
import XLSX from 'xlsx'

function s2ab(s) {
  var buf = new ArrayBuffer(s.length);
  var view = new Uint8Array(buf);
  for (var i = 0; i != s.length; ++i) view[i] = s.charCodeAt(i) & 0xFF;
  return buf;
}
// 通过表格对象拿到表格数据
function generateArray(table) {
  var out = [];
  var rows = table.querySelectorAll('tr');

  for (var R = 0; R < rows.length; ++R) {
    var outRow = [];
    var row = rows[R];
    var columns = row.querySelectorAll('td');
    if (columns.length === 0) {
      columns = row.querySelectorAll('th');
    }
    for (var C = 0; C < columns.length; ++C) {
      var cell = columns[C];
      var cellValue = cell.innerText;
      if (cellValue !== "" && cellValue == +cellValue) cellValue = +cellValue;

      //Handle Value
      outRow.push(cellValue !== "" ? cellValue : null);

      //Handle Colspan
    }
    out.push(outRow);
  }
  return out;
};

export function export_json_to_excel(th, jsonData, defaultTitle) {

  var data = jsonData;
  data.unshift(th);
  var ws_name = "SheetJS";

  const ws = XLSX.utils.aoa_to_sheet(data);
  const wb = XLSX.utils.book_new();

  XLSX.utils.book_append_sheet(wb, ws, ws_name);

  var wbout = XLSX.write(wb, {bookType: 'xlsx', bookSST: false, type: 'binary'});
  var title = defaultTitle || '列表'
  FileSaver.saveAs(new Blob([s2ab(wbout)], {type: "application/octet-stream"}), title + ".xlsx")
}

export function export_table_to_excel(id) {
  var theTable = document.getElementById(id);
  var oo = generateArray(theTable);

  var data = oo;

  var ws_name = "SheetJS";
  const ws = XLSX.utils.aoa_to_sheet(data);
  const wb = XLSX.utils.book_new();


  /* add worksheet to workbook */
  XLSX.utils.book_append_sheet(wb, ws, ws_name);

  var wbout = XLSX.write(wb, {bookType: 'xlsx', bookSST: false, type: 'binary'});

  FileSaver.saveAs(new Blob([s2ab(wbout)], {type: "application/octet-stream"}), "test.xlsx")
}

