"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
var assert_1 = require("assert");
var Docx_1 = require("../classes/Docx");
describe('docx', function () {
    it('Tbl_&_CreateP', function (done) {
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
            var docx = new Docx_1.Docx('test_Tbl_&_CreateP.docx', 'outputUnitTest/');
            docx.createP();
            docx.addContentP('میلاد', { fontColor: 'white', backgroundFont: 'black' });
            docx.createTable(data); // Without Merge
            //docx.createTable(dataMerge); // Merge Table
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
    }); // it Tbl_&_CreateP
    it('Merge_&_CreateP', function (done) {
        try {
            var dataMerge = [
                { x: 1, y: 0, value: '', mergeRow: '2', mergeCol: '3' },
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
            ]; // dataMerge
            var docx = new Docx_1.Docx('test_Merge_&_CreateP.docx', 'outputUnitTest/');
            docx.createP();
            docx.addContentP('میلاد', { fontColor: 'white', backgroundFont: 'red' });
            docx.createTable(dataMerge); // Without Merge
            //docx.createTable(dataMerge); // Merge Table
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
    }); // it Merge_&_CreateP
});
//# sourceMappingURL=test_Merge_&_Tbl_&_CreateP_docx.js.map