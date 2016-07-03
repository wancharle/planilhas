chai.should();

describe('planilhas',function (){
    describe('utilidades',function(){
        it(".json2matrix deve converter json para matrix",function(){
            matrix = planilhas.json2matrix([{a:1,b:2},{a:3,b:4}]);      
            matrix.should.be.deep.equal([['a','b'],[1,2],[3,4]]);
        });
        it(".colname2colnum deve converter A, Z, AA, AB, AZ, BA, ZA, ZZ, AAA  para 0, 25, 26, 27, 51, 52, 676, 701, 702",function(){ 
			result = ['a','Z','Aa','aB','AZ','BA','ZA','ZZ','AAA'].map(planilhas.colname2colnum) 
			console.log('colname2colnum ->', result); 
			result.should.be.deep.equal([0,25,26,27,51,52,676,701,702])
        });
        it(".colnum2colname deve converter 0, 25, 26, 27, 51, 52, 676, 701, 702 para A, Z, AA, AB, AZ, BA, ZA, ZZ, AAA ",function(){
            result = [0,25,26,27,51,52,676,701,702].map(planilhas.colnum2colname)
			console.log('colnum2colname ->', result);
            result.should.be.deep.equal(['A','Z','AA','AB','AZ','BA','ZA','ZZ','AAA'])
        });
        it(".cellname2colrow deve converter A1, B2, AA22, ZZ1, AAA5000 para [0,0], [1,1], [26,21], [701,0], [702,4999]",function(){
            result = ['a1','b2','AA22','zz1','AAA5000'].map(planilhas.cellname2colrow)
			console.log('cellname2colrow ->', result);
            result.should.be.deep.equal([ [0,0], [1,1], [26,21], [701,0], [702,4999] ]);
        });
        it(".colrow2cellname deve converter [0,0], [1,1], [26,21], [701,0], [702,4999] para A1, B2, AA22, ZZ1, AAA5000", function(){
            result = [[0,0], [1,1], [26,21], [701,0], [702,4999] ].map(planilhas.colrow2cellname)
			console.log('colrow2cellname ->', result);
            result.should.be.deep.equal(['A1','B2','AA22','ZZ1','AAA5000'])
        });
        it(".writeCellByName deve aumentar matrix de dados da planilha", function(){
            planilha = { data: [ [1,2], [4,5] ] }
            data = planilhas.writeCellByName(planilha,'c1',3)
            data = planilhas.writeCellByName(planilha,'c2',6)
            data = planilhas.writeCellByName(planilha,'d2',7)
            data = planilhas.writeCellByName(planilha,'b3',8)
            data = planilhas.writeCellByName(planilha,'a4',9)
            console.log('writeCellByName -> ',JSON.stringify(data))
            planilha.data.should.be.deep.equal([[1,2,3],[4,5,6,7],[null,8],[9]]);
        });
        it(".writeRangeByName deve escrever um mesmo dado numa range da planilha", function(){
            planilha = { data: [ [1,2], [4,5] ] }
            data = planilhas.writeRangeByName(planilha,'c2:d4',6)
            data = planilhas.writeRangeByName(planilha,'a4:b4',7)
            console.log('writeRangeByName -> ',JSON.stringify(planilha.data))
            planilha.data.should.be.deep.equal([[1,2],[4,5,6,6],[null,null,6,6],[7,7,6,6]])
        });
        it(".writeCellByName deve escrever style", function(){
            style = { id:1}
            planilha = { data: [ [1,2], [3,4] ] }
            data = planilhas.writeCellByName(planilha,'a3',5,style)
            console.log('writeCellByName -Style -> ',JSON.stringify(planilha.data))
            planilha.data.should.be.deep.equal([[1,2],[3,4],[{ value:5, metadata:{style:1}}]])
        });
        it(".writeCellByName deve escrever style sem alterar value da celula", function(){
            style = { id:2}
            planilha = { data: [ [1,2], [3,4], [7]] }
            data = planilhas.writeCellByName(planilha,'a3',null,style)
            console.log('writeCellByName - samevalue -> ',JSON.stringify(planilha.data))
            planilha.data.should.be.deep.equal([[1,2],[3,4],[{ value:7, metadata:{style:2}}]])
        });
        it(".writeCellByName deve atualizar style sem alterar value da celula", function(){
            style = { id:2}
            planilha = { data: [ [1,2], [3,4], [{value:5,metadata:{style:1}}]] }
            data = planilhas.writeCellByName(planilha,'a3',null,style)
            console.log('writeCellByName - samevalue -> ',JSON.stringify(planilha.data))
            planilha.data.should.be.deep.equal([[1,2],[3,4],[{ value:5, metadata:{style:2}}]])
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
