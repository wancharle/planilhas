// Generated by CoffeeScript 1.10.0
(function() {
  this.planilhas = {};

  if (typeof module === 'object' && module.exports) {
    module.exports = this.planilhas;
  }

}).call(this);
// Generated by CoffeeScript 1.10.0
(function() {
  var XLSX, datenum;

  if (typeof module === 'object' && module.exports) {
    XLSX = require('xlsx');
  } else {
    XLSX = this.XLSX;
  }

  datenum = function(v, date1904) {
    var epoch;
    if (date1904) {
      v += 1462;
    }
    epoch = Date.parse(v);
    return (epoch - new Date(Date.UTC(1899, 11, 30))) / (24 * 60 * 60 * 1000);
  };

  this.planilhas.sheet_from_array_of_arrays = function(data, opts) {
    var C, R, cell, cell_ref, i, j, range, ref, ref1, ws;
    ws = {};
    range = {
      s: {
        c: 10000000,
        r: 10000000
      },
      e: {
        c: 0,
        r: 0
      }
    };
    for (R = i = 0, ref = data.length; 0 <= ref ? i < ref : i > ref; R = 0 <= ref ? ++i : --i) {
      for (C = j = 0, ref1 = data[R].length; 0 <= ref1 ? j < ref1 : j > ref1; C = 0 <= ref1 ? ++j : --j) {
        if (range.s.r > R) {
          range.s.r = R;
        }
        if (range.s.c > C) {
          range.s.c = C;
        }
        if (range.e.r < R) {
          range.e.r = R;
        }
        if (range.e.c < C) {
          range.e.c = C;
        }
        cell = {
          v: data[R][C]
        };
        if (cell.v === null) {
          continue;
        }
        cell_ref = XLSX.utils.encode_cell({
          c: C,
          r: R
        });
        if (typeof cell.v === 'number') {
          cell.t = 'n';
        } else if (typeof cell.v === 'boolean') {
          cell.t = 'b';
        } else if (cell.v instanceof Date) {
          cell.t = 'n';
          cell.z = XLSX.SSF._table[14];
          cell.v = datenum(cell.v);
        } else {
          cell.t = 's';
        }
        ws[cell_ref] = cell;
      }
    }
    if (range.s.c < 10000000) {
      ws['!ref'] = XLSX.utils.encode_range(range);
    }
    return ws;
  };

}).call(this);
// Generated by CoffeeScript 1.10.0
(function() {
  var Workbook, XLSX, buildSheetFromMatrix, fs, s2ab;

  if (typeof module === 'object' && module.exports) {
    XLSX = require('xlsx');
    fs = require('fs');
  } else {
    XLSX = this.XLSX;
  }

  buildSheetFromMatrix = this.planilhas.sheet_from_array_of_arrays;

  Workbook = (function() {
    Workbook.defaults = {
      bookType: 'xlsx',
      bookSST: false,
      type: 'binary'
    };

    function Workbook() {
      this.SheetNames = [];
      this.Sheets = {};
    }

    Workbook.prototype.addSheet = function(data, name, options) {
      if (options == null) {
        options = Workbook.defaults;
      }
      name = name || 'Sheet';
      data = buildSheetFromMatrix(data || [], options);
      this.SheetNames.push(name);
      return this.Sheets[name] = data;
    };

    Workbook.prototype.save = function(options) {
      if (options == null) {
        options = Workbook.defaults;
      }
      return this.excelData = XLSX.write(this, options);
    };

    Workbook.prototype.saveBlob = function(filename) {
      if (filename == null) {
        filename = "test.xlsx";
      }
      return saveAs(new Blob([s2ab(this.excelData)], {
        type: "application/octet-stream"
      }), filename);
    };

    Workbook.prototype.saveFile = function(filename) {
      var buffer, wstream;
      if (filename == null) {
        filename = "test.xlsx";
      }
      buffer = new Buffer(this.excelData, 'binary');
      wstream = fs.createWriteStream(filename);
      wstream.write(buffer);
      return wstream.end();
    };

    return Workbook;

  })();

  s2ab = function(s) {
    var buf, i, j, ref, view;
    buf = new ArrayBuffer(s.length);
    view = new Uint8Array(buf);
    for (i = j = 0, ref = s.length - 1; 0 <= ref ? j < ref : j > ref; i = 0 <= ref ? ++j : --j) {
      view[i] = s.charCodeAt(i) & 0xFF;
    }
    return buf;
  };

  this.planilhas.Workbook = Workbook;

}).call(this);
