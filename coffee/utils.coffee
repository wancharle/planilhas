

if typeof module == 'object' && module.exports
  XLSX = require('xlsx')
else
  XLSX = @XLSX



datenum = (v, date1904) ->
  if (date1904)
    v+=1462
  epoch = Date.parse(v)
  return (epoch - new Date(Date.UTC(1899, 11, 30))) / (24 * 60 * 60 * 1000)


@planilhas.sheet_from_array_of_arrays = (data, opts) ->
  ws = {}
  range = {s: {c:10000000, r:10000000}, e: {c:0, r:0 }}
  for  R in [ 0...(data.length)]
    for C in [ 0...(data[R].length)]
      if range.s.r > R
        range.s.r = R
      if range.s.c > C
        range.s.c = C
      if range.e.r < R
        range.e.r = R
      if range.e.c < C
        range.e.c = C

      cell = {v: data[R][C] }
      if cell.v == null
        continue
      cell_ref = XLSX.utils.encode_cell({c:C,r:R})

      if typeof cell.v == 'number'
        cell.t = 'n'
      else if typeof cell.v == 'boolean'
        cell.t = 'b'
      else if cell.v instanceof Date
        cell.t = 'n'; cell.z = XLSX.SSF._table[14]
        cell.v = datenum(cell.v)
      else
        cell.t = 's'

      ws[cell_ref] = cell
  if range.s.c < 10000000
    ws['!ref'] = XLSX.utils.encode_range(range)
  return ws
 
object2array = (obj,keys)-> keys.map((k)-> obj[k])

@planilhas.json2matrix = (json) ->
  array = []
  if json.length > 0
    keys = Object.keys(json[0])
    array.push(keys)
    for i in [0...json.length]
      obj = object2array(json[i],keys)
      array.push(obj)

  return array
# vim: set ts=2 sw=2 sts=2 expandtab:

