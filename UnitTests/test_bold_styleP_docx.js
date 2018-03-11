"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
var assert_1 = require("assert");
var Docx_1 = require("../classes/Docx");
describe('docx', function () {
    it('bold_styleP_default', function (done) {
        try {
            var docx = new Docx_1.Docx('test_bold_styleP_default.docx', 'outputUnitTest/');
            docx.createP();
            docx.addContentP('میلاد');
            var out = docx.generate();
            if (out == false) {
                assert_1.fail("Don't create File of docx");
            }
            else {
                console.log("Create File of Docx");
                done();
            }
        }
        catch (e) {
            assert_1.fail(e);
            done();
        }
    }); // it test_bold_styleP_docx
    it('bold_styleP_user', function (done) {
        try {
            var docx = new Docx_1.Docx('test_bold_styleP_user.docx', 'outputUnitTest/');
            docx.createP();
            docx.addContentP('میلاد', { bold: 'true' });
            var out = docx.generate();
            if (out == false) {
                assert_1.fail("Don't create File of docx");
            }
            else {
                console.log("Create File of Docx");
                done();
            }
        }
        catch (e) {
            assert_1.fail(e);
            done();
        }
    });
});
//# sourceMappingURL=test_bold_styleP_docx.js.map