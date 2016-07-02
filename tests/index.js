chai.should();

describe('planilhas',function (){
    describe('.json2matrix',function(){
        it("deve converter json para matrix",function(){
            matrix = planilhas.json2matrix([{a:1,b:2},{a:3,b:4}]);      
            matrix.should.be.deep.equal([['a','b'],[1,2],[3,4]]);
        });
    });
    describe('.Workbook',function(){
        it("deve pode criar uma Workbook e salvar em xlsx e mesclar celulas",function(done){
            var wb = new planilhas.Workbook()
            var data = [ ['Hello world', 'lixo'] ]
            wb.addSheet(data,'Hello World!')
            data = [ ['Olá','Mundo!'] ]
            wb.addSheet(data,'Olá Mundo!')

            wb.sheets[0].mergeCells('A1','B1')

            wb.save('helloworld.xlsx')
            done()
        });
        it ("deve poder estilisar e formatar as celulas", function (done){
            var workbook = new planilhas.Workbook()
            // estilos
            var stylesheet = workbook.stylesheet
            var currency = stylesheet.createFormat({ format: '$#,##0.00'});
            var red = 'FFFF0000';
            var importantFormatter = stylesheet.createFormat({
                font: {
                    bold: true,
                    color: red,
                    size: 36
                },
                border: {
                    bottom: {color: red, style: 'thin'},
                    top: {color: red, style: 'thin'},
                    left: {color: red, style: 'thin'},
                    right: {color: red, style: 'thin'}
                }
            });
            var greenBorder = stylesheet.createFormat({
                border: {
                    bottom: {color: 'FF00FFFF', style: 'thin'},
                    top: {color: 'FF00FF00', style: 'thin'},
                }
            });
 
            // dados com estilos
            var data = [
            ['Artist', 'Album', 'Price'],
            ['Buckethead', 'Albino Slug', { value:8.99,metadata: {style: importantFormatter.id }}, { value:0,metadata:{style:greenBorder.id}}, {value:0,metadata:{ style:importantFormatter.id}}],
            ['Buckethead', 'Electric Tears', {value:13.99,metadata:{style: currency.id}}],
            ['Buckethead', 'Colma', 11.34],
            ];
            workbook.addSheet(data,'Estilos e Formatação')
            workbook.sheets[0].mergeCells('C2','e2');
            workbook.save('estilos.xlsx')
            done()
        });
        it ("deve poder adicionar imagens nas planilhas", function (done){
             var wb = new planilhas.Workbook();
             wb.addSheet([],'imagens');
             var medias = new planilhas.Medias(wb.workbook);

             var dimg1 = Q.defer(); 
             medias.addImage('../Imagem1.png','imagem1.png',function (image){
                pic = image.pic            
                pic.createAnchor( 'absoluteAnchor' , {
                     x: planilhas.Positioning.pixelsToEMUs(300),
                     y: planilhas.Positioning.pixelsToEMUs(300),
                     width: planilhas.Positioning.pixelsToEMUs(300),
                     height: planilhas.Positioning.pixelsToEMUs(300)
                });
                dimg1.resolve()
              });
             var dimg2 = Q.defer(); 
             medias.addImage('../Imagem1.png','imagem1.png',function (image){
                pic = image.pic            
                pic.createAnchor('twoCellAnchor', {from: { x: 0,y: 0 }, to: {x: 3,y: 3}});
                dimg2.resolve()
              });
             var dimg3 = Q.defer(); 
             medias.addImage('../Imagem1.png','imagem1.png',function (image){
                pic = image.pic            
                pic.createAnchor('oneCellAnchor', {x: 1,y: 4,width: planilhas.Positioning.pixelsToEMUs(300),height: planilhas.Positioning.pixelsToEMUs(300)});
                dimg3.resolve()
             });


             Q.allSettled([dimg1.promise,dimg2.promise,dimg3.promise]).then(function (){
                medias.drawOn(wb.sheetByName['imagens'])
                wb.save('teste.xlsx');
                done()
             });
        });
    });
});
