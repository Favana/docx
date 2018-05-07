"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
var assert_1 = require("assert");
var Docx_1 = require("../classes/Docx");
describe('style in table', function () {
    it('test_backgroundCel_style_tbl', function (done) {
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
            var style = [
                { x: 1, y: 0 },
                { x: 1, y: 1 },
                { x: 1, y: 2 },
                { x: 2, y: 0, background: 'red' },
                { x: 2, y: 1, background: 'green' },
                { x: 2, y: 2 },
                { x: 3, y: 0, background: 'red' },
                { x: 3, y: 1 },
                { x: 3, y: 2 },
                { x: 4, y: 0 },
                { x: 4, y: 1 },
                { x: 4, y: 2 }
            ]; // style
            var docx = new Docx_1.Docx('test_backgroundCel_style_tbl.docx', 'outputUnitTest/');
            docx.createTable(data, style);
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
    }); //  it test_backgroundCel_style_tbl
}); // describe
//# sourceMappingURL=test_bacgroundCel_style_tbl_docx.js.map