"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
var Docx_1 = require("../classes/Docx");
var assert_1 = require("assert");
describe('docx', function () {
    it('CreateP', function (done) {
        var officegen = new Docx_1.docx('test_createP.docx', 'outputUnitTest/');
        officegen.createP();
        var out = officegen.generate();
        if (out == false) {
            assert_1.fail("Don't create File of docx");
        }
        else {
            console.log("Create File of Docx");
            done();
        }
    }); // it CreateP
    it('addContentPError', function (done) {
        var officegen = new Docx_1.docx('test_createP.docx', 'outputUnitTest/');
        officegen.addContentP('میلاد');
        var out = officegen.generate();
        if (out == false) {
            assert_1.fail("Don't create File of docx");
        }
        else {
            console.log("Create File of Docx");
            done();
        }
    }); // it addContentPError
    it('addContentP', function (done) {
        var officegen = new Docx_1.docx('test_createP.docx', 'outputUnitTest/');
        officegen.createP();
        officegen.addContentP('میلاد');
        var out = officegen.generate();
        if (out == false) {
            assert_1.fail("Don't create File of docx");
        }
        else {
            console.log("Create File of Docx");
            done();
        }
    }); // it addContentP
});
//# sourceMappingURL=test_createP_docx.js.map