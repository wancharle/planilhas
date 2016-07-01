if typeof module == 'object' && module.exports
  XLSX = require('xlsx')
  fs = require('fs')
else
  XLSX = @XLSX

buildSheetFromMatrix = @planilhas.sheet_from_array_of_arrays

class Workbook
  @defaults =
    bookType: 'xlsx',
    bookSST: false,
    type: 'binary'

  constructor: ()->
    @SheetNames = []
    @Sheets = {}
    
  addSheet: (data,name,options = Workbook.defaults)->
      name = name or 'Sheet'
      data = buildSheetFromMatrix(data or [], options)
      @SheetNames.push(name)
      @Sheets[name] = data

  save: (options = Workbook.defaults)->
    @excelData = XLSX.write(@, options)

  saveBlob:(filename="test.xlsx",sheetname="Sheet", charset="iso-8859-1") ->
    if filename.slice(-3).toLowerCase() == "csv"
      console.log('oi')
      csv = XLSX.utils.sheet_to_csv(@Sheets[sheetname])
      console.log(csv)
      saveAs(new Blob([csv],{type:"text/csv;charset="+charset+";"}),filename)
    else
      saveAs(new Blob([s2ab(@excelData)],{type:"application/octet-stream"}),filename)

  saveFile:(filename="test.xlsx") ->
    buffer=new Buffer(@excelData, 'binary')
    wstream = fs.createWriteStream(filename)
    wstream.write(buffer)
    wstream.end()



s2ab = (s) ->
  buf = new ArrayBuffer(s.length)
  view = new Uint8Array(buf)
  for i in [ 0...(s.length-1) ]
    view[i] = s.charCodeAt(i) & 0xFF
  return buf

@planilhas.Workbook = Workbook
# vim: set ts=2 sw=2 sts=2 expandtab:

