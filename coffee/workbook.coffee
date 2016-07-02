ExcelBuilder = @ExcelBuilder

buildSheetFromMatrix = @planilhas.sheet_from_array_of_arrays

class Medias
  constructor: (@workbook)->
    @drawings = new ExcelBuilder.Drawings()
    @workbook.addDrawings(@drawings)

  addMedia: (imagedata,filename,callback)->
    picRef = @workbook.addMedia( 'image', filename , imagedata)
    pic = new ExcelBuilder.Drawing.Picture()
    pic.setMedia(picRef)
    @drawings.addDrawing(pic)
    callback({pic:pic,picRef:picRef})
 
  addImage: (url, filename, callback)->
    self = @
    xhr = new XMLHttpRequest()
    xhr.responseType = 'blob'
    xhr.onload = ->
      reader  = new FileReader()
      reader.onloadend = ->
        imagedata = reader.result.split(',')[1]
        self.addMedia(imagedata,filename,callback)
      reader.readAsDataURL(xhr.response)
    xhr.open('GET', url)
    xhr.send()

  drawOn: (worksheet)->
    worksheet.addDrawings(@drawings)

class Workbook
  constructor: ()->
    @workbook = ExcelBuilder.Builder.createWorkbook()
    @sheets = []
    @sheetByName = {}
    @stylesheet = @workbook.getStyleSheet()
    
  addSheet: (data,name)->
    name = name or "Sheet #{@sheets.length + 1}"
    worksheet = @workbook.createWorksheet({name:name})
    worksheet.setData(data)
    @sheets.push(worksheet)
    @sheetByName[name] = worksheet
    @workbook.addWorksheet(worksheet)

  save: (filename="teste.xlsx")->
    data = planilhas.ExcelBuilder.Builder.createFile(@workbook)
    data.then (dataBase64)=>
      @saveBlob(dataBase64,filename)
      
  saveBlob:(data,filename="test.xlsx") ->
    saveAs(planilhas.base64toBlob(data,"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"),filename)

@planilhas.Medias = Medias
@planilhas.Workbook = Workbook
@planilhas.ExcelBuilder = ExcelBuilder
@planilhas.Positioning = ExcelBuilder.Positioning

# vim: set ts=2 sw=2 sts=2 expandtab:

