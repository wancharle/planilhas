
var rodandoNoBrowser = true;
if (typeof module == 'object' && module.exports){
    var planilhas = require('../planilhas')
    var chai = require('chai')
    rodandoNoBrowser = false
}
chai.should();


describe('planilhas',function (){
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
