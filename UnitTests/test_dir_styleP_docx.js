"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
var assert_1 = require("assert");
var Docx_1 = require("../classes/Docx");
describe('docx', function () {
    it('direction_styleP_default', function (done) {
        try {
            var docx = new Docx_1.Docx('test_dir_styleP_default.docx', 'outputUnitTest/');
            docx.createP();
            docx.addContentP('میلاد');
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
        }
    }); //  it direction_styleP_default
    it('direction_styleP_User', function (done) {
        try {
            var docx = new Docx_1.Docx('test_dir_styleP_user.docx', 'outputUnitTest/');
            docx.createP();
            docx.addContentP('milad', { direction: 'ltr' });
            docx.addContentP('milad', { direction: 'ltr' });
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
    }); // it direction_styleP_user
});
//# sourceMappingURL=test_dir_styleP_docx.js.map