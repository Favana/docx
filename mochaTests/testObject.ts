import {} from 'mocha';
import {createWriteStream} from 'fs'
let officegen = require('officegen');


describe('testOfficegen', function(){
    let fileName = 'test.docx';
    let filePath = 'outpotProject/'
    it('docx', function(done){
        let docx = officegen('docx');
         let objP =  docx.createP();
         objP.addText('علی خانه است');
         objP.addText('رضا بیرون نرفت');
        let out = createWriteStream(filePath+fileName);
        docx.generate(out);
        done();
    });


});