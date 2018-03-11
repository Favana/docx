"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
var assert_1 = require("assert");
var Docx_1 = require("../classes/Docx");
describe('docx', function () {
    it('fontSize_StyleP_User', function (done) {
        try {
            var docx = new Docx_1.Docx('test_fontSize_styleP_User.docx', 'outputUnitTest/');
            docx.createP();
            docx.addContentP('میلاد', { fontSize: 30 });
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
    }); // it fontSize_StyleP_User
    it('fontSize_StyleP_default', function (done) {
        try {
            var docx = new Docx_1.Docx('test_fontSize_styleP_default.docx', 'outputUnitTest/');
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
    });
});
//# sourceMappingURL=test_fontSize_styleP_docx.js.map