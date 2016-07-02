 
object2array = (obj,keys)-> keys.map((k)-> obj[k])


# string byte to arraybuffer
s2ab = (s) ->
  buf = new ArrayBuffer(s.length)
  view = new Uint8Array(buf)
  for i in [ 0...(s.length-1) ]
    view[i] = s.charCodeAt(i) & 0xFF
  return buf

@planilhas.base64toBlob = (base64data, contentType)-> new Blob([s2ab(atob(base64data))],{type:contentType})
  
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

