"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
var fs_1 = require("fs");
var officegen = require('officegen');
describe('testOfficegen', function () {
    var fileName = 'test.docx';
    var filePath = 'outpotProject/';
    it('docx', function (done) {
        var docx = officegen('docx');
        var objP = docx.createP();
        objP.addText('علی خانه است');
        objP.addText('رضا بیرون نرفت');
        var out = fs_1.createWriteStream(filePath + fileName);
        docx.generate(out);
        done();
    });
});
//# sourceMappingURL=testObject.js.map