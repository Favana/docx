import {} from 'mocha';
import {createWriteStream} from 'fs'
let officegen = require('officegen');


describe('testOfficegen', function(){
    let fileName = 'test.docx';
    let filePath = 'outpotProject/'
    it('docx', function(done){
        let docx = officegen('docx');
        // let objP =  docx.createP();
        // objP.addText()
        docx.addText('milad');
        let out = createWriteStream(fileName+fileName);
        docx.generate(out)
    });


});