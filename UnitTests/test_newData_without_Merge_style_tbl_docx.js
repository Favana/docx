"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
var assert_1 = require("assert");
var Docx_1 = require("../classes/Docx");
var _ = require('lodash');
describe('style in table with new data', function () {
    it('style Without merge', function (done) {
        var data = [
            { x: 1, y: 0, value: '', mergeRow: '', mergeCol: '', style: { borderSize: 30 } },
            { x: 1, y: 1, value: 'سال 1390', mergeRow: '', mergeCol: '' },
            { x: 1, y: 2, value: 'سال1391', mergeRow: '', mergeCol: '' },
            { x: 2, y: 0, value: 'کل', mergeRow: '', mergeCol: '', style: { borderSize: 30 } },
            { x: 2, y: 1, value: '21545288', mergeRow: '', mergeCol: '' },
            { x: 2, y: 2, value: '85487525', mergeRow: '', mergeCol: '' },
            { x: 3, y: 0, value: 'البرز', mergeRow: '', mergeCol: '', style: { fontFamily: 'B Nazanin', align: 'right' } },
            { x: 3, y: 1, value: '2521', mergeRow: '', mergeCol: '' },
            { x: 3, y: 2, value: '5485', mergeRow: '', mergeCol: '' },
            { x: 4, y: 0, value: 'بندرعباس', mergeRow: '', mergeCol: '' },
            { x: 4, y: 1, value: '145214', mergeRow: '', mergeCol: '' },
            { x: 4, y: 2, value: '2255', mergeRow: '', mergeCol: '' }
        ]; // data
        var docx = new Docx_1.Docx('test_newData_without_style_tbl.docx', 'outputUnitTest/');
        docx.createTable(data);
        var out = docx.generate();
        if (out == false) {
            assert_1.fail("Don't create File of docx");
        }
        else {
            console.log("Create File of Docx");
            done();
        }
    }); //  it for style Without merge
}); // describe
//# sourceMappingURL=test_newData_without_Merge_style_tbl_docx.js.map