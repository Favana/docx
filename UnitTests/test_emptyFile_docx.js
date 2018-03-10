"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
var Docx_1 = require("../classes/Docx");
var assert_1 = require("assert");
describe('docx', function () {
    it('generate Docx File', function (done) {
        try {
            var officegen = new Docx_1.docx('test_emptyFile.docx', 'outputUnitTest/');
            var out = officegen.generate();
            if (out == false) {
                assert_1.fail("Don't create File of docx");
            }
            else {
                console.log('Create File of Docx');
                done();
            }
        }
        catch (e) {
            assert_1.fail(e);
            done();
        }
    }); // it generate Docx File
    it('generate Docx File Error for null input', function (done) {
        try {
            var officegen = new Docx_1.docx('', ''); //fileName is null OR filePath is null
            var out = officegen.generate();
            if (out == false) {
                assert_1.fail("Don't create File of docx");
            }
            else {
                console.log('Create File of Docx');
                done();
            }
        }
        catch (e) {
            if (e == "Don't Null inputs docx function OR Don't send parameters for docx function") {
                console.log(e);
                done();
            }
            else {
                assert_1.fail(e);
            }
        }
    }); // it generate Docx File Error for null input
    it('generate Docx File Error for the mistake input', function (done) {
        try {
            var officegen = new Docx_1.docx('outputUnitTest/', 'test_emptyFile.docx');
            var out = officegen.generate();
            if (out == false) {
                assert_1.fail("Don't create File of docx");
            }
            else {
                console.log('Create File of Docx');
                done();
            }
        }
        catch (e) {
            if (e == "Don't Null inputs docx function OR Don't send parameters for docx function") {
                console.log(e);
                done();
            }
            else {
                assert_1.fail(e);
            }
        }
    }); // it generate Docx File Error for the mistake input
    it('generate Docx File Error for the filePath is null input', function (done) {
        try {
            var officegen = new Docx_1.docx('test_emptyFile.docx', '');
            var out = officegen.generate();
            if (out == false) {
                assert_1.fail("Don't create File of docx");
            }
            else {
                console.log('Create File of Docx');
                done();
            }
        }
        catch (e) {
            if (e == "Don't Null inputs docx function OR Don't send parameters for docx function") {
                console.log(e);
                done();
            }
            else {
                assert_1.fail(e);
                done();
            }
        }
    }); // it generate Docx File Error for the filePath is null input
    it('generate Docx File Error for the fileName is null input', function (done) {
        try {
            var officegen = new Docx_1.docx('', 'outputUnitTest/');
            var out = officegen.generate();
            if (out == false) {
                assert_1.fail("Don't create File of docx");
            }
            else {
                console.log('Create File of Docx');
                done();
            }
        }
        catch (e) {
            if (e == "Don't Null inputs docx function OR Don't send parameters for docx function") {
                console.log(e);
                done();
            }
            else {
                assert_1.fail(e);
                done();
            }
        }
    });
}); // describe
//# sourceMappingURL=test_emptyFile_docx.js.map