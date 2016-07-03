 
object2array = (obj,keys)-> keys.map((k)-> obj[k])


# string byte to arraybuffer
s2ab = (s) ->
  buf = new ArrayBuffer(s.length)
  view = new Uint8Array(buf)
  for i in [ 0...(s.length-1) ]
    view[i] = s.charCodeAt(i) & 0xFF
  return buf

@planilhas.base64toBlob = (base64data, contentType)->
  new Blob([s2ab(atob(base64data))],{type:contentType})
  
@planilhas.json2matrix = (json) ->
  array = []
  if json.length > 0
    keys = Object.keys(json[0])
    array.push(keys)
    for i in [0...json.length]
      obj = object2array(json[i],keys)
      array.push(obj)

  return array

colname2colnum = (name) ->
  base = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'
  result = 0
  j = name.length
  for c, i in name.toUpperCase()
      j-=1
      result+=(base.indexOf(c)+1)*Math.pow(base.length,j)
  return result - 1
@planilhas.colname2colnum = colname2colnum

colnum2colname = (n) ->
  ordA = 'A'.charCodeAt(0)
  ordZ = 'Z'.charCodeAt(0)
  len = ordZ - ordA + 1

  s = ""
  while n >= 0
    s = String.fromCharCode(n % len + ordA) + s
    n = Math.floor(n / len) - 1
  return s
@planilhas.colnum2colname = colnum2colname

splitCellname = (cellname) -> cellname.match(/[a-zA-Z]+|[0-9]+/g)
@planilhas.splitCellname = splitCellname

cellname2colrow = (cellname) ->
  [colname, row] = splitCellname(cellname)
  col = colname2colnum(colname)
  row = row - 1
  return [col, row]
@planilhas.cellname2colrow = cellname2colrow

colrow2cellname = (colrow) ->
  [col,row] = colrow
  colname = colnum2colname(col)
  return "#{colname}#{row + 1}"
@planilhas.colrow2cellname = colrow2cellname

sheetAccomodateCell = (sheet,col,row)->
  data = sheet.data
  # ajusta rows
  rowsNeeded = (row + 1) - data.length
  if rowsNeeded > 0
    data.push [] for i in [1..rowsNeeded]
  # ajusta cols
  colsNeeded = (col + 1) - data[row].length
  if colsNeeded > 0
    data[row].push(null)  for i in [1...colsNeeded]
  return sheet.data
  
sheetAccomodateRange = (sheet,col1,row1,col2,row2)->
  rowMax = if row1 > row2 then row1 else row2
  rowMin = if row1 < row2 then row1 else row2
  col = if col1 > col2 then col1 else col2
  sheetAccomodateCell(sheet,col,row) for row in [rowMin..rowMax]
  return sheet.data

writeCell = (sheet,col,row,value,style) ->
  sheetAccomodateCell(sheet,col,row)
  if style
    if value == null
      # aplica/altera style em celula existente sem alterar o conteudo
      value = sheet.data[row][col]
      if typeof value == 'object'
        value.metadata.style =  style.id
      else
        value = { value: value, metadata: { style: style.id }}
    else
      value = { value: value, metadata: { style: style.id }}
  sheet.data[row][col] = value
  return sheet.data
@planilhas.writeCell = writeCell

writeCellByName = (sheet,cellname,value,style) ->
  [col, row] = cellname2colrow(cellname)
  writeCell(sheet,col,row,value,style)
@planilhas.writeCellByName = writeCellByName

writeRange = (sheet,col1,row1,col2,row2,value,style) ->
  for row in [row1..row2]
    for col in [col1..col2]
      writeCell(sheet,col,row,value,style)
  return sheet.data
@planilhas.writeRange = writeRange

writeRangeByName = (sheet,rangename,value,style) ->
  [cellname1,cellname2] = rangename.split(":")
  [col1,row1] = cellname2colrow(cellname1)
  [col2,row2] = cellname2colrow(cellname2)
  writeRange(sheet,col1,row1,col2,row2,value,style)
  
@planilhas.writeRangeByName = writeRangeByName
  
    
     

# vim: set ts=2 sw=2 sts=2 expandtab:
