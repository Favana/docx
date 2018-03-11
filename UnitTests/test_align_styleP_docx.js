"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
var assert_1 = require("assert");
var Docx_1 = require("../classes/Docx");
describe('docx', function () {
    it('test_align_styleP_default(right)', function (done) {
        try {
            var docx = new Docx_1.Docx('test_align_styleP_defualt.docx', 'outputUnitTest/');
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
    }); // it test_align_styleP_default(right)
    it('test_align_styleP_user_leftAlign', function (done) {
        try {
            var docx = new Docx_1.Docx('test_align_styleP_user_left.docx', 'outputUnitTest/');
            docx.createP();
            docx.addContentP('میلاد', { align: 'left' });
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
    }); // it test_align_styleP_user_leftAlign
    it('test_align_styleP_user_centerAlign', function (done) {
        try {
            var docx = new Docx_1.Docx('test_align_styleP_user_center.docx', 'outputUnitTest/');
            docx.createP();
            docx.addContentP('میلاد', { align: 'center' });
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
    }); // it test_align_styleP_user_centerAlign
});
//# sourceMappingURL=test_align_styleP_docx.js.map