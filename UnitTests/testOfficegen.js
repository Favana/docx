"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
var officegen = require("officegen");
var fs_1 = require("fs");
describe('testOfficegen', function () {
    it('docxTable', function (done) {
        var docx = officegen('docx');
        var table = [
            [{
                    val: "من اليمين",
                    opts: {
                        b: true,
                        sz: '48',
                        shd: {
                            fill: "7F7F7F",
                            themeFill: "text1",
                            "themeFillTint": "80"
                        },
                        fontFamily: "Avenir Book",
                        rtl: true,
                    }
                }, {
                    val: "Title1",
                    opts: {
                        b: true,
                        color: "A00000",
                        align: "right",
                        shd: {
                            fill: "92CDDC",
                            themeFill: "text1",
                            "themeFillTint": "80"
                        }
                    }
                }]
        ];
        var tableStyle = {
            borders: false,
            rtl: true
        };
        docx.createTable(table, tableStyle);
        docx.createTable(table, tableStyle);
        var out = fs_1.createWriteStream('outputUnitTest/' + 'testOfficegen.docx');
        docx.generate(out);
        console.log("create Docx File");
        done();
    });
});
//# sourceMappingURL=testOfficegen.js.map