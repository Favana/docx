"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
var Docx_1 = require("../classes/Docx");
var assert_1 = require("assert");
describe('docx', function () {
    it('Create File Docx', function (done) {
        var docxOffice = new Docx_1.docx('test_emptyFile.docx', 'outputUnitTest/');
        var out = docxOffice.generate();
        if (out == false) {
            assert_1.fail("Don't create File of docx");
        }
        else {
            console.log('Create File of Docx');
            done();
        }
    });
}); // describe
//# sourceMappingURL=test_emptyFile_docx.js.map