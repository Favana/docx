"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
var Docx_1 = require("../classes/Docx");
var assert_1 = require("assert");
describe('docx', function () {
    it('CreateP', function (done) {
        try {
            var docx = new Docx_1.Docx('test_createP.docx', 'outputUnitTest/');
            docx.createP();
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
    }); // it CreateP
    it('addContentPError', function (done) {
        try {
            var docx = new Docx_1.Docx('test_createP.docx', 'outputUnitTest/');
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
            if (e == 'createP function is undefined') {
                console.log(e);
                done();
            }
            else {
                assert_1.fail(e);
                done();
            }
        }
    }); // it addContentPError
    it('addContentP', function (done) {
        try {
            var docx = new Docx_1.Docx('test_createP.docx', 'outputUnitTest/');
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
    }); // it addContentP
    it('font_StyleP_Error For mistake send first parameter(text)', function (done) {
        try {
            var docx = new Docx_1.Docx('test_createP.docx', 'outputUnitTest/');
            docx.createP();
            docx.addContentP({ fontFamily: 'B Nazanin' });
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
    });
});
//# sourceMappingURL=test_createP_docx.js.map