"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
var assert_1 = require("assert");
var Docx_1 = require("../classes/Docx");
describe('docx', function () {
    it('font_StyleP_User', function (done) {
        var officegen = new Docx_1.docx('test_font_styleP_User.docx', 'outputUnitTest/');
        officegen.createP();
        officegen.addContentP('میلاد', { fontFamily: 'B Elham' });
        var out = officegen.generate();
        if (out == false) {
            assert_1.fail("Don't create File of docx");
        }
        else {
            console.log("Create File of Docx");
            done();
        }
    }); // it for font_StyleP_User
    it('font_StyleP_default', function (done) {
        var officegen = new Docx_1.docx('test_font_styleP.docx', 'outputUnitTest/');
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
    }); // it for font_StyleP_default
});
//# sourceMappingURL=test_font_styleP_docx.js.map