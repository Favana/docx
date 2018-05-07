"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
var json2xml = require('json2xml');
var _ = require('lodash');
var table = /** @class */ (function () {
    function table() {
    }
    table.prototype.createTr = function (_body, data) {
        _body;
        var check = _body['w:tbl'].length;
        if (check == 1) {
            _body;
        }
        else {
            for (var i = 1; i < check; i++) {
                _body['w:tbl'].pop();
            }
            //this.globalTbl = (<any>Object).assign(resultTable, this.globalTbl);
            // let p = {'w:p':[]};
            // _body = (<any>Object).assign(p, _body);
            _body;
            //  return _body;
        }
        _body;
        var counterRow = _.uniqBy(data, 'x');
        var row = counterRow.length;
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
                            { 'w:r': [
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
    table.prototype.tableStyle = function (_body, style) {
        style;
        _body;
        if (style != null) {
            if (typeof style == 'object') {
                var counterX = _.uniqBy(style, 'x');
                var counterY = _.uniqBy(style, 'y');
                var counterCol = counterY.length;
                var counterRow = counterX.length;
                for (var i = 1; i <= counterRow; i++) {
                    for (var j = 0; j < counterCol; j++) {
                        var find = _.find(style, { x: i, y: j });
                        var borderSize = find.borderSize;
                        var align = find.align;
                        var fontFamily = find.fontFamily;
                        var fontColor = find.fontColor;
                        var fontSize = find.fontSize;
                        var backgroundCell = find.background;
                        var bold = find.bold;
                        if (borderSize != undefined) { ///// if check sizeBorder
                            _body['w:tbl'][i]['w:tr'][j]['w:tc'][0]['w:tcPr'].push({ 'w:tblBorders': [
                                    { 'w:top': '', attr: { 'w:val': 'single', 'w:sz': borderSize, 'w:space': '0', 'w:color': 'auto' } },
                                    { 'w:left': '', attr: { 'w:val': 'single', 'w:sz': borderSize, 'w:space': '0', 'w:color': 'auto' } },
                                    { 'w:bottom': '', attr: { 'w:val': 'single', 'w:sz': borderSize, 'w:space': '0', 'w:color': 'auto' } },
                                    { 'w:right': '', attr: { 'w:val': 'single', 'w:sz': borderSize, 'w:space': '0', 'w:color': 'auto' } }
                                ] } // 'w:tblBorders'
                            );
                            _body;
                        } // if check sizeBorder
                        if (align != undefined) { // if check align
                            var check = _body['w:tbl'][i]['w:tr'][j]['w:tc'][1]['w:p'].length;
                            if (check == 1) {
                                _body['w:tbl'][i]['w:tr'][j]['w:tc'][1]['w:p'].splice(0, 0, { 'w:pPr': [
                                        { 'w:jc': '', attr: { 'w:val': align } }
                                    ] });
                                _body;
                            }
                            else if (check > 1) {
                                _body['w:tbl'][i]['w:tr'][j]['w:tc'][1]['w:p'][0]['w:pPr'].push({ 'w:jc': '', attr: { 'w:val': align } });
                                _body;
                            }
                        } // if check align
                        if (fontFamily != undefined) { // if  check font
                            var checkP = _body['w:tbl'][i]['w:tr'][j]['w:tc'][1]['w:p'].length;
                            if (checkP == 1) {
                                var checkR = _body['w:tbl'][i]['w:tr'][j]['w:tc'][1]['w:p'][0]['w:r'].length;
                                if (checkR == 1) {
                                    _body['w:tbl'][i]['w:tr'][j]['w:tc'][1]['w:p'][0]['w:r'].splice(0, 0, { 'w:rPr': [
                                            { 'w:rFonts': '', attr: { 'w:cs': fontFamily } },
                                            { 'w:rtl': '' }
                                        ] } // 'w:rPr'
                                    );
                                    _body;
                                }
                                else if (checkR > 1) {
                                    _body['w:tbl'][i]['w:tr'][j]['w:tc'][1]['w:p'][0]['w:r'][0]['w:rPr'].push({ 'w:rFonts': '', attr: { 'w:cs': fontFamily } }, { 'w:rtl': '' });
                                    _body;
                                }
                            }
                            else if (checkP > 1) {
                                var checkR = _body['w:tbl'][i]['w:tr'][j]['w:tc'][1]['w:p'][1]['w:r'].length;
                                if (checkR == 1) {
                                    _body['w:tbl'][i]['w:tr'][j]['w:tc'][1]['w:p'][1]['w:r'].splice(0, 0, { 'w:rPr': [
                                            { 'w:rFonts': '', attr: { 'w:cs': fontFamily } },
                                            { 'w:rtl': '' }
                                        ] } // 'w:rPr'
                                    );
                                    _body;
                                }
                                else if (checkR > 1) {
                                    _body['w:tbl'][i]['w:tr'][j]['w:tc'][1]['w:p'][1]['w:r'][0]['w:rPr'].push({ 'w:rFonts': '', attr: { 'w:cs': fontFamily } }, { 'w:rtl': '' });
                                    _body;
                                }
                            }
                        } // if  check font
                        if (fontColor != undefined) { // if  fontColor
                            var checkP = _body['w:tbl'][i]['w:tr'][j]['w:tc'][1]['w:p'].length;
                            if (checkP == 1) {
                                var checkR = _body['w:tbl'][i]['w:tr'][j]['w:tc'][1]['w:p'][0]['w:r'].length;
                                if (checkR == 1) {
                                    _body['w:tbl'][i]['w:tr'][j]['w:tc'][1]['w:p'][0]['w:r'].splice(0, 0, { 'w:rPr': [
                                            { 'w:color': '', attr: { 'w:val': fontColor } }
                                        ] } // 'w:rPr'
                                    );
                                    _body;
                                }
                                else if (checkR > 1) {
                                    _body['w:tbl'][i]['w:tr'][j]['w:tc'][1]['w:p'][0]['w:r'][0]['w:rPr'].push({ 'w:color': '', attr: { 'w:val': fontColor } });
                                    _body;
                                }
                            }
                            else if (checkP > 1) {
                                var checkR = _body['w:tbl'][i]['w:tr'][j]['w:tc'][1]['w:p'][1]['w:r'].length;
                                if (checkR == 1) {
                                    _body['w:tbl'][i]['w:tr'][j]['w:tc'][1]['w:p'][1]['w:r'].splice(0, 0, { 'w:rPr': [
                                            { 'w:color': '', attr: { 'w:val': fontColor } }
                                        ] } // 'w:rPr'
                                    );
                                    _body;
                                }
                                else if (checkR > 1) {
                                    _body['w:tbl'][i]['w:tr'][j]['w:tc'][1]['w:p'][1]['w:r'][0]['w:rPr'].push({ 'w:color': '', attr: { 'w:val': fontColor } });
                                    _body;
                                }
                            }
                        } // if check fontColor
                        if (fontSize != undefined) { // if check fontSize
                            var checkP = _body['w:tbl'][i]['w:tr'][j]['w:tc'][1]['w:p'].length;
                            fontSize = 2 * fontSize;
                            if (checkP == 1) {
                                var checkR = _body['w:tbl'][i]['w:tr'][j]['w:tc'][1]['w:p'][0]['w:r'].length;
                                if (checkR == 1) {
                                    _body['w:tbl'][i]['w:tr'][j]['w:tc'][1]['w:p'][0]['w:r'].splice(0, 0, { 'w:rPr': [
                                            { 'w:sz': '', attr: { 'w:val': fontSize } }
                                        ] } // 'w:rPr'
                                    );
                                    _body;
                                }
                                else if (checkR > 1) {
                                    _body['w:tbl'][i]['w:tr'][j]['w:tc'][1]['w:p'][0]['w:r'][0]['w:rPr'].push({ 'w:sz': '', attr: { 'w:val': fontSize } });
                                    _body;
                                }
                            }
                            else if (checkP > 1) {
                                var checkR = _body['w:tbl'][i]['w:tr'][j]['w:tc'][1]['w:p'][1]['w:r'].length;
                                if (checkR == 1) {
                                    _body['w:tbl'][i]['w:tr'][j]['w:tc'][1]['w:p'][1]['w:r'].splice(0, 0, { 'w:rPr': [
                                            { 'w:sz': '', attr: { 'w:val': fontSize } }
                                        ] } // 'w:rPr'
                                    );
                                    _body;
                                }
                                else if (checkR > 1) {
                                    _body['w:tbl'][i]['w:tr'][j]['w:tc'][1]['w:p'][1]['w:r'][0]['w:rPr'].push({ 'w:sz': '', attr: { 'w:val': fontSize } });
                                    _body;
                                }
                            }
                        } //  if check fontSize
                        if (backgroundCell != undefined) { // if check background cell table
                            _body['w:tbl'][i]['w:tr'][j]['w:tc'][0]['w:tcPr'].push({ 'w:shd': '', attr: { 'w:val': 'clear', 'w:color': 'auto', 'w:fill': backgroundCell } });
                            _body;
                        } // if check background cell table
                        if (bold != undefined) { // if for check bold
                            var checkP = _body['w:tbl'][i]['w:tr'][j]['w:tc'][1]['w:p'].length;
                            if (checkP == 1) {
                                var checkR = _body['w:tbl'][i]['w:tr'][j]['w:tc'][1]['w:p'][0]['w:r'].length;
                                if (checkR == 1) {
                                    _body['w:tbl'][i]['w:tr'][j]['w:tc'][1]['w:p'][0]['w:r'].splice(0, 0, { 'w:rPr': [
                                            { 'w:bCs': '' } // <w:bCs/>
                                        ] } // 'w:rPr'
                                    );
                                    _body;
                                }
                                else if (checkR > 1) {
                                    _body['w:tbl'][i]['w:tr'][j]['w:tc'][1]['w:p'][0]['w:r'].push({ 'w:bCs': '' } // <w:bCs/>
                                    );
                                    _body;
                                }
                            }
                            else if (checkP > 1) {
                                var checkR = _body['w:tbl'][i]['w:tr'][j]['w:tc'][1]['w:p'][1]['w:r'].length;
                                if (checkR == 1) {
                                    _body['w:tbl'][i]['w:tr'][j]['w:tc'][1]['w:p'][1]['w:r'].splice(0, 0, { 'w:rPr': [
                                            { 'w:bCs': '' } // <w:bCs/>
                                        ] } // 'w:rPr'
                                    );
                                    _body;
                                }
                                else if (checkR > 1) {
                                    _body['w:tbl'][i]['w:tr'][j]['w:tc'][1]['w:p'][1]['w:r'].push({ 'w:bCs': '' } // <w:bCs/>
                                    );
                                    _body;
                                }
                            }
                        } // if for check bold
                    } // for j
                } // for i
                _body;
            }
            else {
                throw "Type of style is not object";
            } // else
        }
        else {
            return _body;
        } //  else
    };
    table.prototype.callingMethod = function (body, data, style) {
        var objTr = this.createTr(body, data);
        var objTc = this.createTc(objTr, data);
        var objP = this.createPtable(objTc, data);
        var objMerge = this.createMerge(objP, data);
        var objStyle = this.tableStyle(objMerge, style);
        return objMerge;
        // next(json2xml(objMerge,{ attributes_key:'attr' } ));
    };
    return table;
}()); //  class
exports.table = table;
//# sourceMappingURL=createTable.js.map