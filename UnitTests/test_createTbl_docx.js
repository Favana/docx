"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
var assert_1 = require("assert");
var Docx_1 = require("../classes/Docx");
describe('docx', function () {
    it('createTbl_docx', function (done) {
        try {
            var data = [
                { x: 1, y: 0, value: '', mergeRow: '', mergeCol: '' },
                { x: 1, y: 1, value: 'سال 1390', mergeRow: '', mergeCol: '' },
                { x: 1, y: 2, value: 'سال1391', mergeRow: '', mergeCol: '' },
                { x: 2, y: 0, value: 'کل', mergeRow: '', mergeCol: '' },
                { x: 2, y: 1, value: '21545288', mergeRow: '', mergeCol: '' },
                { x: 2, y: 2, value: '85487525', mergeRow: '', mergeCol: '' },
                { x: 3, y: 0, value: 'البرز', mergeRow: '', mergeCol: '' },
                { x: 3, y: 1, value: '2521', mergeRow: '', mergeCol: '' },
                { x: 3, y: 2, value: '5485', mergeRow: '', mergeCol: '' },
                { x: 4, y: 0, value: 'بندرعباس', mergeRow: '', mergeCol: '' },
                { x: 4, y: 1, value: '145214', mergeRow: '', mergeCol: '' },
                { x: 4, y: 2, value: '2255', mergeRow: '', mergeCol: '' }
            ]; // data
            var docx = new Docx_1.Docx('test_createTbl.docx', 'outputUnitTest/');
            docx.createTable(data);
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
    }); // it createTbl_docx
    it('createTbl_Error For  mistake send first parameter(data)', function (done) {
        try {
            var docx = new Docx_1.Docx('test_createTbl.docx', 'outputUnitTest/');
            docx.createTable('milad');
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
            if (e == 'The first parameter sent is wrong') {
                console.log(e);
                done();
            }
            else {
                assert_1.fail(e);
                done();
            }
        }
    }); // it  For  mistake And Empty send first parameter(data)
    it('createTbl_Error For  Empty send first parameter(data)', function (done) {
        try {
            var docx = new Docx_1.Docx('test_createTbl.docx', 'outputUnitTest/');
            docx.createTable();
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
            if (e == 'The first parameter sent is wrong') {
                console.log(e);
                done();
            }
            else {
                assert_1.fail(e);
                done();
            }
        }
    }); // it  For  Empty send first parameter(data)
});
//# sourceMappingURL=test_createTbl_docx.js.map