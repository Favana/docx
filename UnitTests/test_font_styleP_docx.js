"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
var assert_1 = require("assert");
var Docx_1 = require("../classes/Docx");
describe('docx', function () {
    it('font_StyleP_User', function (done) {
        try {
            var docx = new Docx_1.Docx('test_font_styleP_User.docx', 'outputUnitTest/');
            docx.createP();
            docx.addContentP('میلاد', { fontFamily: 'B Elham' });
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
    }); // it for font_StyleP_User
    it('font_StyleP_default', function (done) {
        try {
            var docx = new Docx_1.Docx('test_font_styleP_default.docx', 'outputUnitTest/');
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
    }); // it for font_StyleP_default
    it('font_StyleP_Error For mistake send first parameter(text)', function (done) {
        try {
            var docx = new Docx_1.Docx('test_font_styleP_User.docx', 'outputUnitTest/');
            docx.createP();
            docx.addContentP({ fontFamily: 'B Elham' });
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
            if (e == 'The sending parameter is incorrect') {
                console.log(e);
                done();
            }
            else if (e == 'createP function is undefined') {
                console.log(e);
                done();
            }
            else {
                assert_1.fail(e);
                done();
            }
        }
    }); //  it font_StyleP_Error For mistake send first parameter
    it('font_StyleP_Error For mistake send second parameter(object)', function (done) {
        try {
            var docx = new Docx_1.Docx('test_font_styleP_User.docx', 'outputUnitTest/');
            docx.createP();
            docx.addContentP('میلاد', 'error');
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
            if (e == 'The sending parameter is incorrect') {
                console.log(e);
                done();
            }
            else if (e == 'createP function is undefined') {
                console.log(e);
                done();
            }
            else {
                assert_1.fail(e);
                done();
            }
        }
    }); // it font_StyleP_Error For mistake send second parameter
});
//# sourceMappingURL=test_font_styleP_docx.js.map