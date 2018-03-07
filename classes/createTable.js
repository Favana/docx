"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
var json2xml = require('json2xml');
var _ = require('lodash');
var table = /** @class */ (function () {
    function table() {
    }
    table.prototype.createTr = function (_body, data) {
        var counterRow = _.uniqBy(data, 'x');
        var row = counterRow.length;
        // let _body = {'w:tbl':[
        //     {'w:tblPr':[ //<w:tblBorders>
        //             {'w:tblBorders':[ //<w:top w:val="single" w:sz="12" w:space="0" w:color="95B3D7" w:themeColor="accent1" w:themeTint="99"/>
        //                     {'w:top':'', attr:{'w:val':"single",  'w:sz':"12",'w:color':'95B3D7', 'w:themeColor':'accent1', 'w:themeTint':'99'}}
        //              ]}
        //     ]}
        //     ]};
        for (var i = 0; i < row; i++) {
            _body['w:tbl'].push({ 'w:tr': [] });
        }
        return _body;
    }; //createTr
    table.prototype.createTc = function (_body, data) {
        var counterCol = _.uniqBy(data, 'y');
        var counterRow = _.uniqBy(data, 'x');
        var row = counterRow.length;
        var col = counterCol.length;
        for (var i = 1; i <= row; i++) {
            for (var j = 0; j < col; j++) {
                _body['w:tbl'][i]['w:tr'].push({ 'w:tc': [{ 'w:tcPr': [{ 'w:tcW': '', attr: { 'w:w': 4657, 'w:type': 'dxa' } }] },] });
            } // for column
        } //for row
        return _body;
    }; //createTc
    table.prototype.createPtable = function (_body, data) {
        var defaultStyle = {
            fontFamily: 'B Nazanin',
            fontSize: 20,
            fontColor: 'black',
            bold: 'false',
            direction: 'rtl',
            align: 'right',
            backgroundFont: 'black'
        };
        var counterCol = _.uniqBy(data, 'y');
        var counterRow = _.uniqBy(data, 'x');
        var row = counterRow.length;
        var col = counterCol.length;
        for (var i = 1; i <= row; i++) {
            for (var j = 0; j < col; j++) {
                var checkValue = _.find(data, { x: i, y: j });
                var value = checkValue.value;
                if (value == '') {
                    _body['w:tbl'][i]['w:tr'][j]['w:tc'].push({ 'w:p': [] });
                }
                else {
                    _body['w:tbl'][i]['w:tr'][j]['w:tc'].push({ 'w:p': [
                            { 'w:pPr': [
                                    { 'w:jc': '', attr: { 'w:val': 'center' } }
                                ] },
                            { 'w:r': [
                                    { 'w:rPr': [
                                            { 'w:rtl': [] },
                                            { 'w:rFonts': '', attr: { 'w:cs': defaultStyle.fontFamily, 'w:hint': 'cs' } },
                                            { 'w:szCs': '', attr: { 'w:val': 2 * (defaultStyle.fontSize) } },
                                            { 'w:color': '', attr: { 'w:val': defaultStyle.fontColor } },
                                        ] },
                                    { 'w:t': value }
                                ] }
                        ] });
                }
            } // for col
        } // for row
        return _body;
    }; //createP
    table.prototype.createMerge = function (_body, data) {
        var counterCol = _.uniqBy(data, 'y');
        var counterRow = _.uniqBy(data, 'x');
        var row = counterRow.length;
        var col = counterCol.length;
        var start = [];
        for (var i = 1; i <= row; i++) {
            for (var j = 0; j < col; j++) {
                var checkMerge = _.find(data, { x: i, y: j });
                var rowX = checkMerge.mergeRow;
                var colY = checkMerge.mergeCol;
                if (colY != '' && rowX != '') {
                    var myX = checkMerge.x;
                    var myY = checkMerge.y;
                    var counterY = colY; // mergeCol
                    var counterX = rowX; //mergeRow
                    /// start merge Row and Col
                    _body['w:tbl'][myX]['w:tr'][myY]['w:tc'][0]['w:tcPr'].push({ 'w:vMerge': '', attr: { 'w:val': 'restart' } }, { 'w:gridSpan': '', attr: { 'w:val': colY } }); ///<w:vMerge w:val="restart"/>
                    var b = _body;
                    ///// loop for merge Col:
                    for (var i_1 = 1; i_1 < counterY; i_1++) {
                        myY = myY + 1;
                        delete _body['w:tbl'][myX]['w:tr'][myY]; // merge col
                    } // for
                    /// end loop for
                    var c = _body;
                    ///// loop for merge Row:
                    for (var i_2 = 1; i_2 < counterX; i_2++) {
                        myY = 0;
                        myY = checkMerge.y;
                        myX = myX + 1;
                        _body['w:tbl'][myX]['w:tr'][myY]['w:tc'][0]['w:tcPr'].push({ 'w:vMerge': [] }, { 'w:gridSpan': '', attr: { 'w:val': colY } }); // merge row
                    } // for
                    /// end loop for
                    var z = _body;
                    ///// loop for merge Col:
                    for (var i_3 = 1; i_3 < counterY; i_3++) {
                        myY = myY + 1;
                        delete _body['w:tbl'][myX]['w:tr'][myY]; // merge col
                    } // for
                    /// end loop for
                    var a = _body;
                } // if
                else if (colY != '' && rowX == '') {
                    var myX = checkMerge.x;
                    var myY = checkMerge.y;
                    var counterY = colY;
                    /// start merge Row and Col
                    _body['w:tbl'][myX]['w:tr'][myY]['w:tc'][0]['w:tcPr'].push({ 'w:gridSpan': '', attr: { 'w:val': colY } }); ///<w:gridSpan w:val="2"/>
                    var b = _body;
                    ///// loop for merge Col:
                    for (var i_4 = 1; i_4 < counterY; i_4++) {
                        myY = myY + 1;
                        delete _body['w:tbl'][myX]['w:tr'][myY]; // merge col
                    } // for
                    var c = _body;
                } // else if
                else if (colY == '' && rowX != '') {
                    var myY = checkMerge.y;
                    var myX = checkMerge.x;
                    var counterX = rowX;
                    /// start merge Row and Row
                    _body['w:tbl'][myX]['w:tr'][myY]['w:tc'][0]['w:tcPr'].push({ 'w:vMerge': '', attr: { 'w:val': 'restart' } }); ///<w:vMerge w:val="restart"/>
                    var b = _body;
                    ///// loop for merge Row:
                    for (var i_5 = 1; i_5 < counterX; i_5++) {
                        myX = myX + 1;
                        _body['w:tbl'][myX]['w:tr'][myY]['w:tc'][0]['w:tcPr'].push({ 'w:vMerge': [] }); // merge row
                    } // for
                    var c = _body;
                } // else ifs
            } // for col
        } // for row
        return _body;
    }; //createMerge
    table.prototype.callingMethod = function (body, data) {
        var objTr = this.createTr(body, data);
        var objTc = this.createTc(objTr, data);
        var objP = this.createPtable(objTc, data);
        var objMerge = this.createMerge(objP, data);
        return objMerge;
        // next(json2xml(objMerge,{ attributes_key:'attr' } ));
    };
    return table;
}()); //  class
exports.table = table;
//# sourceMappingURL=createTable.js.map