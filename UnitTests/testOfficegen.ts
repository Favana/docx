import * as officegen from 'officegen';
import {} from 'mocha';
import {fail} from 'assert';
import {createWriteStream} from  'fs'

describe('testOfficegen', function(){
    it('docxTable', function(done){
        let docx = officegen('docx');
        let table = [
            [{
                val: "من اليمين",
                opts: {
                    b:true,
                    sz: '48',
                    shd: {
                        fill: "7F7F7F",
                        themeFill: "text1",
                        "themeFillTint": "80"
                    },
                    fontFamily: "Avenir Book",
                    rtl: true,
                }
            },{
                val: "Title1",
                opts: {
                    b:true,
                    color: "A00000",
                    align: "right",
                    shd: {
                        fill: "92CDDC",
                        themeFill: "text1",
                        "themeFillTint": "80"
                    }
                }
            }]];
        let tableStyle = {
            borders: false,
            rtl: true
        };
            docx.createTable(table, tableStyle);
        docx.createTable(table, tableStyle);
        let out = createWriteStream('outputUnitTest/' + 'testOfficegen.docx');
        docx.generate(out);
        console.log("create Docx File");
        done()



    });
});