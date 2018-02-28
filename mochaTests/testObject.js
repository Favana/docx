"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
var fs_1 = require("fs");
var officegen = require('officegen');
describe('testOfficegen', function () {
    var fileName = 'test.docx';
    var filePath = 'outpotProject/';
    it('docx', function (done) {
        var docx = officegen('docx');
        // let objP =  docx.createP();
        // objP.addText()
        docx.addText('milad');
        var out = fs_1.createWriteStream(fileName + fileName);
        docx.generate(out);
    });
});
//# sourceMappingURL=testObject.js.map