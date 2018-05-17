"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
var Docx_1 = require("./classes/Docx");
/****
 @ObjectStyle

 {
      fontFamily : 'value',
      fontSize : value,
      fontColor : 'value',
      bold: 'true',
      direction :'default rtl',
      align : 'value',
      background: 'value',
      borderSize: value
  }
 *****/
var data = [
    { x: 1, y: 0, value: 'استان', mergeRow: '3', mergeCol: '', style: { background: '#DCDCDC', align: 'center', fontFamily: 'B Nazanin' } },
    { x: 1, y: 1, value: 'طول شرقی', mergeRow: '', mergeCol: '4', style: { background: '#DCDCDC', align: 'center', fontFamily: 'B Nazanin' } },
    { x: 1, y: 2, value: '', mergeRow: '', mergeCol: '' },
    { x: 1, y: 3, value: '', mergeRow: '', mergeCol: '' },
    { x: 1, y: 4, value: '', mergeRow: '', mergeCol: '' },
    { x: 2, y: 0, value: '', mergeRow: '', mergeCol: '' },
    { x: 2, y: 1, value: 'حداقل', mergeRow: '', mergeCol: '2', style: { background: '#DCDCDC', align: 'center', fontFamily: 'B Nazanin' } },
    { x: 2, y: 2, value: '', mergeRow: '', mergeCol: '' },
    { x: 2, y: 3, value: 'حداکثر', mergeRow: '', mergeCol: '2', style: { background: '#DCDCDC', align: 'center', fontFamily: 'B Nazanin' } },
    { x: 2, y: 4, value: '', mergeRow: '', mergeCol: '' },
    { x: 3, y: 0, value: '', mergeRow: '', mergeCol: '' },
    { x: 3, y: 1, value: 'دقیقه', mergeRow: '', mergeCol: '', style: { background: '#DCDCDC', align: 'center', fontFamily: 'B Nazanin' } },
    { x: 3, y: 2, value: 'درجه', mergeRow: '', mergeCol: '', style: { background: '#DCDCDC', align: 'center', fontFamily: 'B Nazanin' } },
    { x: 3, y: 3, value: 'دقیقه', mergeRow: '', mergeCol: '', style: { background: '#DCDCDC', align: 'center', fontFamily: 'B Nazanin' } },
    { x: 3, y: 4, value: 'درجه', mergeRow: '', mergeCol: '', style: { background: '#DCDCDC', align: 'center', fontFamily: 'B Nazanin' } },
    { x: 4, y: 0, value: 'کل کشور', mergeRow: '', mergeCol: '', style: { background: '#DCDCDC', align: 'center', fontFamily: 'B Nazanin' } },
    { x: 4, y: 1, value: '4', mergeRow: '', mergeCol: '', style: { align: 'center', fontFamily: 'B Nazanin' } },
    { x: 4, y: 2, value: '2', mergeRow: '', mergeCol: '', style: { align: 'center', fontFamily: 'B Nazanin' } },
    { x: 4, y: 3, value: '145214', mergeRow: '', mergeCol: '', style: { align: 'center', fontFamily: 'B Nazanin' } },
    { x: 4, y: 4, value: '145214', mergeRow: '', mergeCol: '', style: { align: 'center', fontFamily: 'B Nazanin' } },
    { x: 5, y: 0, value: 'تهران', mergeRow: '', mergeCol: '', style: { background: '#DCDCDC', align: 'center', fontFamily: 'B Nazanin' } },
    { x: 5, y: 1, value: '145214', mergeRow: '', mergeCol: '', style: { align: 'center', fontFamily: 'B Nazanin' } },
    { x: 5, y: 2, value: '145214', mergeRow: '', mergeCol: '', style: { align: 'center', fontFamily: 'B Nazanin' } },
    { x: 5, y: 3, value: '145214', mergeRow: '', mergeCol: '', style: { align: 'center', fontFamily: 'B Nazanin' } },
    { x: 5, y: 4, value: '145214', mergeRow: '', mergeCol: '', style: { align: 'center', fontFamily: 'B Nazanin' } },
    { x: 6, y: 0, value: 'گیلان', mergeRow: '', mergeCol: '', style: { background: '#DCDCDC', align: 'center', fontFamily: 'B Nazanin' } },
    { x: 6, y: 1, value: '145214', mergeRow: '', mergeCol: '', style: { align: 'center', fontFamily: 'B Nazanin' } },
    { x: 6, y: 2, value: '145214', mergeRow: '', mergeCol: '', style: { align: 'center', fontFamily: 'B Nazanin' } },
    { x: 6, y: 3, value: '145214', mergeRow: '', mergeCol: '', style: { align: 'center', fontFamily: 'B Nazanin' } },
    { x: 6, y: 4, value: '145214', mergeRow: '', mergeCol: '', style: { align: 'center', fontFamily: 'B Nazanin' } },
    { x: 7, y: 0, value: 'مازندران', mergeRow: '', mergeCol: '', style: { background: '#DCDCDC', align: 'center', fontFamily: 'B Nazanin' } },
    { x: 7, y: 1, value: '145214', mergeRow: '', mergeCol: '', style: { align: 'center', fontFamily: 'B Nazanin' } },
    { x: 7, y: 2, value: '145214', mergeRow: '', mergeCol: '', style: { align: 'center', fontFamily: 'B Nazanin' } },
    { x: 7, y: 3, value: '145214', mergeRow: '', mergeCol: '', style: { align: 'center', fontFamily: 'B Nazanin' } },
    { x: 7, y: 4, value: '145214', mergeRow: '', mergeCol: '', style: { align: 'center', fontFamily: 'B Nazanin' } },
    { x: 8, y: 0, value: 'البرز', mergeRow: '', mergeCol: '', style: { background: '#DCDCDC', align: 'center', fontFamily: 'B Nazanin' } },
    { x: 8, y: 1, value: '145214', mergeRow: '', mergeCol: '', style: { align: 'center', fontFamily: 'B Nazanin' } },
    { x: 8, y: 2, value: '145214', mergeRow: '', mergeCol: '', style: { align: 'center', fontFamily: 'B Nazanin' } },
    { x: 8, y: 3, value: '145214', mergeRow: '', mergeCol: '', style: { align: 'center', fontFamily: 'B Nazanin' } },
    { x: 8, y: 4, value: '145214', mergeRow: '', mergeCol: '', style: { align: 'center', fontFamily: 'B Nazanin' } }
]; // data
var docx = new Docx_1.Docx('test_index.docx', 'output/');
docx.createTable(data);
var out = docx.generate();
if (out == false) {
    console.log("Don't create File of docx");
}
else {
    console.log("Create File of Docx");
}
//# sourceMappingURL=index.js.map