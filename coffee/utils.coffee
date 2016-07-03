 
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

@planilhas.colname2colnum = (name) ->
  base = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'
  result = 0
  j = name.length
  for c, i in name.toUpperCase()
      j-=1
      result+=(base.indexOf(c)+1)*Math.pow(base.length,j)
  return result - 1

@planilhas.colnum2colname = (n) ->
  ordA = 'A'.charCodeAt(0)
  ordZ = 'Z'.charCodeAt(0)
  len = ordZ - ordA + 1

  s = ""
  while n >= 0
    s = String.fromCharCode(n % len + ordA) + s
    n = Math.floor(n / len) - 1
  return s

@planilhas.splitCellname = (cellname) -> cellname.match(/[a-zA-Z]+|[0-9]+/g)

@planilhas.cellname2colrow = (cellname) ->
  [colname, row] = @planilhas.splitCellname(cellname)
  col = @planilhas.colname2colnum(colname)
  row = row - 1
  return [col, row]

@planilhas.colrow2cellname = (colrow) ->
  [col,row] = colrow
  colname = @planilhas.colnum2colname(col)
  return "#{colname}#{row + 1}"

# vim: set ts=2 sw=2 sts=2 expandtab:
