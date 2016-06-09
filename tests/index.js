
var rodandoNoBrowser = true;
if (typeof module == 'object' && module.exports){
    var planilhas = require('../planilhas')
    var chai = require('chai')
    rodandoNoBrowser = false
}
chai.should();


describe('planilhas',function (){
    describe('.json2matrix',function(){
        it("deve converter json para matrix",function(){
            matrix = planilhas.json2matrix([{a:1,b:2},{a:3,b:4}]);      
            matrix.should.be.deep.equal([['a','b'],[1,2],[3,4]]);
        });
    });
    describe('.Workbook',function(){
        it("deve pode criar uma planilha Workbook",function(){
            wb = new planilhas.Workbook()
            wb.should.have.property('Sheets');
            wb.should.have.property('SheetNames');
        });
        it ("deve poder salvar a planilha", function (){
             var wb = new planilhas.Workbook();
             var sheet = [ ['a','b', 'c'],[1,2],[3,4],[5,6] ] 
             wb.addSheet(sheet)
             wb.save()
             if (rodandoNoBrowser)
                 wb.saveBlob()
             else
                 wb.saveFile()
             wb.should.have.property('excelData');
     
        });

    });



});
