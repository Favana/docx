"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
var json2xml = require('json2xml');
var _ = require('lodash');
var table = /** @class */ (function () {
    function table() {
        /**@
         *
         * @type {string}
         * private globalTable ===>  Table Body Object
         * private globalData  ===>  The data sent from the server side
         */
        this.globalTable = " ";
        this.globalData = " ";
    }
    table.prototype.createTr = function (_body, data) {
        this.globalData = data;
        var check = _body['w:tbl'].length;
        if (check == 1) {
            _body;
        }
        else {
            for (var i = 1; i < check; i++) {
                _body['w:tbl'].pop();
            }
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
                _body['w:tbl'][i]['w:tr'].push({
                    'w:tc': [{
                            'w:tcPr': [{
                                    'w:tcW': '',
                                    attr: { 'w:w': 4657, 'w:type': 'dxa' }
                                }]
                        },]
                });
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
                    _body['w:tbl'][i]['w:tr'][j]['w:tc'].push({
                        'w:p': [
                            {
                                'w:r': [
                                    { 'w:t': value }
                                ]
                            }
                        ]
                    });
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
                    myY;
                    var counterY = colY; // mergeCol
                    var counterX = rowX; //mergeRow
                    /// start merge Row and Col
                    _body['w:tbl'][myX]['w:tr'][myY]['w:tc'][0]['w:tcPr'].push({ 'w:vMerge': '', attr: { 'w:val': 'restart' } }, { 'w:gridSpan': '', attr: { 'w:val': colY } }); //
                    var b = _body;
                    b;
                    var newWidth = colY * 4657;
                    newWidth;
                    var test_1 = _body['w:tbl'][myX]['w:tr'][myY]['w:tc'][0]['w:tcPr'].splice(0, 1, {
                        'w:tcW': '', attr: { 'w:w': newWidth, 'w:type': 'dxa' }
                    });
                    ///// loop for merge Col:
                    for (var i_1 = 1; i_1 < counterY; i_1++) {
                        myY = myY + 1;
                        delete _body['w:tbl'][myX]['w:tr'][myY]; // merge col
                    } // for
                    /// end loop for
                    var c = _body;
                    c;
                    ///// loop for merge Row:
                    for (var i_2 = 1; i_2 < counterX; i_2++) {
                        myY = 0;
                        myY = checkMerge.y;
                        myX = myX + 1;
                        _body['w:tbl'][myX]['w:tr'][myY]['w:tc'][0]['w:tcPr'].push({ 'w:vMerge': [] }, {
                            'w:gridSpan': '',
                            attr: { 'w:val': colY }
                        }); // merge row
                    } // for
                    /// end loop for
                    var z = _body;
                    z;
                    ///// loop for merge Col:
                    for (var i_3 = 1; i_3 < counterY; i_3++) {
                        myY = myY + 1;
                        myX;
                        delete _body['w:tbl'][myX]['w:tr'][myY]; // merge col
                    } // for
                    /// end loop for
                    var a = _body;
                    a;
                } // if
                else if (colY != '' && rowX == '') {
                    var myX = checkMerge.x;
                    var myY = checkMerge.y;
                    var counterY = colY;
                    counterY;
                    /// start merge Row and Col
                    _body['w:tbl'][myX]['w:tr'][myY]['w:tc'][0]['w:tcPr'].push({
                        'w:gridSpan': '',
                        attr: { 'w:val': colY }
                    }); ///<w:gridSpan w:val="2"/>
                    var newWidth = colY * 4657;
                    newWidth;
                    var test_2 = _body['w:tbl'][myX]['w:tr'][myY]['w:tc'][0]['w:tcPr'].splice(0, 1, {
                        'w:tcW': '', attr: { 'w:w': newWidth, 'w:type': 'dxa' }
                    });
                    ///// loop for merge Col:
                    for (var i_4 = 1; i_4 < counterY; i_4++) {
                        myY = myY + 1;
                        delete _body['w:tbl'][myX]['w:tr'][myY]; // merge col
                    } // for
                    var c = _body;
                    c;
                } // else if
                else if (colY == '' && rowX != '') {
                    var myY = checkMerge.y;
                    var myX = checkMerge.x;
                    var counterX = rowX;
                    /// start merge Row and Row
                    _body['w:tbl'][myX]['w:tr'][myY]['w:tc'][0]['w:tcPr'].push({
                        'w:vMerge': '',
                        attr: { 'w:val': 'restart' }
                    }); ///<w:vMerge w:val="restart"/>
                    var b = _body;
                    b;
                    ///// loop for merge Row:
                    for (var i_5 = 1; i_5 < counterX; i_5++) {
                        myX = myX + 1;
                        _body['w:tbl'][myX]['w:tr'][myY]['w:tc'][0]['w:tcPr'].push({ 'w:vMerge': [] });
                    } // for
                    var c = _body;
                    c;
                } // else ifs
            } // for col
        } // for row
        this.globalTable = _body;
    }; //createMerge
    table.prototype.tableStyle = function () {
        this.globalTable; // tbl body object
        this.globalData; // The data sent from the server side
        var count = this.globalData.length;
        var styleObject = {};
        var styleData = [];
        for (var i = 0; i < count; i++) {
            var infoData = _.find(this.globalData, this.globalData[i]);
            var x = infoData.x;
            var y = infoData.y;
            var checkMergeCol = infoData.mergeCol;
            var style = infoData.style;
            var check = Object.keys(styleObject).length;
            if (check == 0) {
                styleObject = { x: x, y: y, style: style };
                styleData.push(styleObject);
            }
            else if (check > 0) {
                if (style == undefined) {
                    styleObject = { x: x, y: y };
                    styleData.push(styleObject);
                }
                else {
                    styleObject = { x: x, y: y, style: style };
                    styleData.push(styleObject);
                }
            }
        } //  for
        if (styleData != null) {
            if (typeof styleData == 'object') {
                var counterX = _.uniqBy(styleData, 'x');
                var counterY = _.uniqBy(styleData, 'y');
                var counterCol = counterY.length;
                var counterRow = counterX.length;
                for (var i = 1; i <= counterRow; i++) {
                    for (var j = 0; j < counterCol; j++) {
                        var find = _.find(styleData, { x: i, y: j });
                        var newStyle = find.style;
                        if (newStyle != undefined) {
                            var topBorderSize = newStyle.topBorderSize;
                            var bottomBorderSize = newStyle.bottomBorderSize;
                            var leftBorderSize = newStyle.leftBorderSize;
                            var rightBorderSize = newStyle.rightBorderSize;
                            var align = newStyle.align;
                            var fontFamily = newStyle.fontFamily;
                            var fontColor = newStyle.fontColor;
                            var fontSize = newStyle.fontSize;
                            var backgroundCell = newStyle.background;
                            var bold = newStyle.bold;
                            var topBorderColor = newStyle.topBorderColor;
                            var rightBorderColor = newStyle.rightBorderColor;
                            var leftBorderColor = newStyle.leftBorderColor;
                            var bottomBorderColor = newStyle.bottomBorderColor;
                            var topBorder = newStyle.topBorder;
                            var bottomBorder = newStyle.bottomBorder;
                            var leftBorder = newStyle.leftBorder;
                            var rightBorder = newStyle.rightBorder;
                            ///////////////////////////////////////////////////CHECK BORDER  ///////////////////////////////////////////////
                            if (topBorder != undefined) {
                                this.globalTable['w:tbl'][i]['w:tr'][j]['w:tc'][0]['w:tcPr'].push({
                                    'w:tcBorders': [
                                        { 'w:top': '', attr: { 'w:val': 'nil' } }
                                    ]
                                }); // push
                            } //  check borderTop
                            if (bottomBorder != undefined) {
                                this.globalTable['w:tbl'][i]['w:tr'][j]['w:tc'][0]['w:tcPr'].push({
                                    'w:tcBorders': [
                                        { 'w:bottom': '', attr: { 'w:val': 'nil' } }
                                    ]
                                }); // push
                                this.globalTable;
                            } // check borderBottom
                            if (rightBorder != undefined) {
                                this.globalTable['w:tbl'][i]['w:tr'][j]['w:tc'][0]['w:tcPr'].push({
                                    'w:tcBorders': [
                                        { 'w:left': '', attr: { 'w:val': 'nil' } }
                                    ]
                                }); // push
                            } //  check border Right
                            if (leftBorder != undefined) {
                                this.globalTable['w:tbl'][i]['w:tr'][j]['w:tc'][0]['w:tcPr'].push({
                                    'w:tcBorders': [
                                        { 'w:right': '', attr: { 'w:val': 'nil' } }
                                    ]
                                }); // push
                            } //  check border Left
                            this.globalTable;
                            /////////////////////////////////////////////////// END ///////////////////////////////////////////////
                            /////////////////////////////////////////////////// BORDER SIZE ///////////////////////////////////////////////
                            if (topBorder != undefined) {
                                this.globalTable;
                            }
                            else {
                                if (topBorderSize != undefined) { ///// if check topBorderSize
                                    this.globalTable['w:tbl'][i]['w:tr'][j]['w:tc'][0]['w:tcPr'].push({
                                        'w:tblBorders': [
                                            {
                                                'w:top': '',
                                                attr: {
                                                    'w:val': 'single',
                                                    'w:sz': topBorderSize,
                                                    'w:space': '0',
                                                }
                                            }
                                        ]
                                    }); // push
                                } // if check topBorderSize
                            } // check topBorder active
                            if (bottomBorder != undefined) {
                                this.globalTable;
                            }
                            else {
                                if (bottomBorderSize != undefined) { // if check bottomBorderSize
                                    this.globalTable['w:tbl'][i]['w:tr'][j]['w:tc'][0]['w:tcPr'].push({
                                        'w:tblBorders': [
                                            {
                                                'w:bottom': '',
                                                attr: {
                                                    'w:val': 'single',
                                                    'w:sz': bottomBorderSize,
                                                    'w:space': '0',
                                                }
                                            }
                                        ]
                                    }); // push
                                } //  if check bottomBorderSize
                            } // if check  bottomBorder Active
                            if (rightBorder) {
                                this.globalTable;
                            }
                            else {
                                if (rightBorderSize != undefined) { // if check rightBorderSize
                                    this.globalTable['w:tbl'][i]['w:tr'][j]['w:tc'][0]['w:tcPr'].push({
                                        'w:tblBorders': [
                                            {
                                                'w:left': '',
                                                attr: {
                                                    'w:val': 'single',
                                                    'w:sz': rightBorderSize,
                                                    'w:space': '0',
                                                }
                                            }
                                        ]
                                    }); // push
                                } //  if check rightBorderSize
                            }
                            if (leftBorder != undefined) {
                                this.globalTable;
                            }
                            else {
                                if (leftBorderSize != undefined) { // if check rightBorderSize
                                    this.globalTable['w:tbl'][i]['w:tr'][j]['w:tc'][0]['w:tcPr'].push({
                                        'w:tblBorders': [
                                            {
                                                'w:right': '',
                                                attr: {
                                                    'w:val': 'single',
                                                    'w:sz': leftBorderSize,
                                                    'w:space': '0',
                                                }
                                            }
                                        ]
                                    }); // push
                                } //  if check rightBorderSize
                            }
                            this.globalTable;
                            ///////////////////////////////////////////////////End ///////////////////////////////////////////////
                            // //////////////////////////////////////////// BORDER  COLOR//////////////////////////////////////////////////
                            if (topBorderColor != undefined) { //// if check for topBorderColor
                                if (topBorder = !undefined) {
                                    this.globalTable;
                                }
                                else {
                                    if (topBorderSize != undefined) {
                                        this.globalTable['w:tbl'][i]['w:tr'][j]['w:tc'][0]['w:tcPr'].push({
                                            'w:tblBorders': [
                                                { 'w:top': '', attr: { 'w:val': 'single', 'w:sz': topBorderSize, 'w:space': '0', 'w:color': topBorderColor } }
                                            ]
                                        }); // push
                                    }
                                    else {
                                        this.globalTable['w:tbl'][i]['w:tr'][j]['w:tc'][0]['w:tcPr'].push({
                                            'w:tblBorders': [
                                                { 'w:top': '', attr: { 'w:val': 'single', 'w:space': '0', 'w:color': topBorderColor } }
                                            ]
                                        }); // push
                                    } //  else
                                }
                            } //  if check for topBorderColor
                            //
                            if (bottomBorderColor != undefined) { //  if check for borderBottomColor
                                if (bottomBorder != undefined) {
                                    this.globalTable;
                                }
                                else {
                                    if (bottomBorderSize != undefined) {
                                        this.globalTable['w:tbl'][i]['w:tr'][j]['w:tc'][0]['w:tcPr'].push({
                                            'w:tblBorders': [
                                                { 'w:bottom': '', attr: { 'w:val': 'single', 'w:sz': bottomBorderSize, 'w:space': '0', 'w:color': bottomBorderColor } }
                                            ]
                                        }); // push
                                    }
                                    else {
                                        this.globalTable['w:tbl'][i]['w:tr'][j]['w:tc'][0]['w:tcPr'].push({
                                            'w:tblBorders': [
                                                { 'w:bottom': '', attr: { 'w:val': 'single', 'w:space': '0', 'w:color': bottomBorderColor } }
                                            ]
                                        }); // push
                                    } //  else
                                }
                            } //  if check for borderBottomColor
                            //
                            if (rightBorderColor != undefined) { //if check for borderRightColor
                                if (rightBorder != undefined) {
                                    this.globalTable;
                                }
                                else {
                                    if (rightBorderSize != undefined) {
                                        this.globalTable['w:tbl'][i]['w:tr'][j]['w:tc'][0]['w:tcPr'].push({
                                            'w:tblBorders': [
                                                { 'w:left': '', attr: { 'w:val': 'single', 'w:sz': rightBorderSize, 'w:space': '0', 'w:color': rightBorderColor } }
                                            ]
                                        }); // push
                                    }
                                    else {
                                        this.globalTable['w:tbl'][i]['w:tr'][j]['w:tc'][0]['w:tcPr'].push({
                                            'w:tblBorders': [
                                                { 'w:left': '', attr: { 'w:val': 'single', 'w:space': '0', 'w:color': rightBorderColor } }
                                            ]
                                        }); // push
                                    } //  else
                                }
                            } //if check for borderRightColor
                            //
                            if (leftBorderColor != undefined) { //if check for borderLeftColor
                                if (leftBorder != undefined) {
                                    this.globalTable;
                                }
                                else {
                                    if (leftBorderSize != undefined) {
                                        this.globalTable['w:tbl'][i]['w:tr'][j]['w:tc'][0]['w:tcPr'].push({
                                            'w:tblBorders': [
                                                { 'w:right': '', attr: { 'w:val': 'single', 'w:sz': leftBorderSize, 'w:space': '0', 'w:color': leftBorderColor } }
                                            ]
                                        }); // push
                                    }
                                    else {
                                        this.globalTable['w:tbl'][i]['w:tr'][j]['w:tc'][0]['w:tcPr'].push({
                                            'w:tblBorders': [
                                                { 'w:right': '', attr: { 'w:val': 'single', 'w:space': '0', 'w:color': leftBorderColor } }
                                            ]
                                        }); // push
                                    } //  else
                                } //if check for borderLeftColor
                                //
                            } //  check for active leftBorder
                            // ////////////////////////////////////////////END///////////////////////////////////////////////////
                            if (align != undefined) { // if check align
                                var check = this.globalTable['w:tbl'][i]['w:tr'][j]['w:tc'][1]['w:p'].length;
                                if (check == 1) {
                                    this.globalTable['w:tbl'][i]['w:tr'][j]['w:tc'][1]['w:p'].splice(0, 0, {
                                        'w:pPr': [
                                            { 'w:jc': '', attr: { 'w:val': align } }
                                        ]
                                    });
                                    this.globalTable['w:tbl'][i]['w:tr'][j]['w:tc'][0]['w:tcPr'].push({ 'w:vAlign': '', attr: { 'w:val': align } });
                                }
                                else if (check > 1) {
                                    this.globalTable['w:tbl'][i]['w:tr'][j]['w:tc'][1]['w:p'][0]['w:pPr'].push({ 'w:jc': '', attr: { 'w:val': align } });
                                    this.globalTable['w:tbl'][i]['w:tr'][j]['w:tc'][0]['w:tcPr'].push({ 'w:vAlign': '', attr: { 'w:val': align } });
                                }
                            } // if check align
                            if (fontFamily != undefined) { // if  check font
                                var checkP = this.globalTable['w:tbl'][i]['w:tr'][j]['w:tc'][1]['w:p'].length;
                                if (checkP == 1) {
                                    var checkR = this.globalTable['w:tbl'][i]['w:tr'][j]['w:tc'][1]['w:p'][0]['w:r'].length;
                                    if (checkR == 1) {
                                        this.globalTable['w:tbl'][i]['w:tr'][j]['w:tc'][1]['w:p'][0]['w:r'].splice(0, 0, {
                                            'w:rPr': [
                                                { 'w:rFonts': '', attr: { 'w:cs': fontFamily } },
                                                { 'w:rtl': '' }
                                            ]
                                        } // 'w:rPr'
                                        );
                                    }
                                    else if (checkR > 1) {
                                        this.globalTable['w:tbl'][i]['w:tr'][j]['w:tc'][1]['w:p'][0]['w:r'][0]['w:rPr'].push({ 'w:rFonts': '', attr: { 'w:cs': fontFamily } }, { 'w:rtl': '' });
                                    }
                                }
                                else if (checkP > 1) {
                                    var checkR = this.globalTable['w:tbl'][i]['w:tr'][j]['w:tc'][1]['w:p'][1]['w:r'].length;
                                    if (checkR == 1) {
                                        this.globalTable['w:tbl'][i]['w:tr'][j]['w:tc'][1]['w:p'][1]['w:r'].splice(0, 0, {
                                            'w:rPr': [
                                                { 'w:rFonts': '', attr: { 'w:cs': fontFamily } },
                                                { 'w:rtl': '' }
                                            ]
                                        } // 'w:rPr'
                                        );
                                    }
                                    else if (checkR > 1) {
                                        this.globalTable['w:tbl'][i]['w:tr'][j]['w:tc'][1]['w:p'][1]['w:r'][0]['w:rPr'].push({ 'w:rFonts': '', attr: { 'w:cs': fontFamily } }, { 'w:rtl': '' });
                                    }
                                }
                            } // if  check font
                            if (fontColor != undefined) { // if  fontColor
                                var checkP = this.globalTable['w:tbl'][i]['w:tr'][j]['w:tc'][1]['w:p'].length;
                                if (checkP == 1) {
                                    var checkR = this.globalTable['w:tbl'][i]['w:tr'][j]['w:tc'][1]['w:p'][0]['w:r'].length;
                                    if (checkR == 1) {
                                        this.globalTable['w:tbl'][i]['w:tr'][j]['w:tc'][1]['w:p'][0]['w:r'].splice(0, 0, {
                                            'w:rPr': [
                                                { 'w:color': '', attr: { 'w:val': fontColor } }
                                            ]
                                        } // 'w:rPr'
                                        );
                                    }
                                    else if (checkR > 1) {
                                        this.globalTable['w:tbl'][i]['w:tr'][j]['w:tc'][1]['w:p'][0]['w:r'][0]['w:rPr'].push({ 'w:color': '', attr: { 'w:val': fontColor } });
                                    }
                                }
                                else if (checkP > 1) {
                                    var checkR = this.globalTable['w:tbl'][i]['w:tr'][j]['w:tc'][1]['w:p'][1]['w:r'].length;
                                    if (checkR == 1) {
                                        this.globalTable['w:tbl'][i]['w:tr'][j]['w:tc'][1]['w:p'][1]['w:r'].splice(0, 0, {
                                            'w:rPr': [
                                                { 'w:color': '', attr: { 'w:val': fontColor } }
                                            ]
                                        } // 'w:rPr'
                                        );
                                    }
                                    else if (checkR > 1) {
                                        this.globalTable['w:tbl'][i]['w:tr'][j]['w:tc'][1]['w:p'][1]['w:r'][0]['w:rPr'].push({ 'w:color': '', attr: { 'w:val': fontColor } });
                                    }
                                }
                            } // if check fontColor
                            if (fontSize != undefined) { // if check fontSize
                                var checkP = this.globalTable['w:tbl'][i]['w:tr'][j]['w:tc'][1]['w:p'].length;
                                fontSize = 2 * fontSize;
                                if (checkP == 1) {
                                    var checkR = this.globalTable['w:tbl'][i]['w:tr'][j]['w:tc'][1]['w:p'][0]['w:r'].length;
                                    if (checkR == 1) {
                                        this.globalTable['w:tbl'][i]['w:tr'][j]['w:tc'][1]['w:p'][0]['w:r'].splice(0, 0, {
                                            'w:rPr': [
                                                { 'w:sz': '', attr: { 'w:val': fontSize } }
                                            ]
                                        } // 'w:rPr'
                                        );
                                    }
                                    else if (checkR > 1) {
                                        this.globalTable['w:tbl'][i]['w:tr'][j]['w:tc'][1]['w:p'][0]['w:r'][0]['w:rPr'].push({ 'w:sz': '', attr: { 'w:val': fontSize } });
                                    }
                                }
                                else if (checkP > 1) {
                                    var checkR = this.globalTable['w:tbl'][i]['w:tr'][j]['w:tc'][1]['w:p'][1]['w:r'].length;
                                    if (checkR == 1) {
                                        this.globalTable['w:tbl'][i]['w:tr'][j]['w:tc'][1]['w:p'][1]['w:r'].splice(0, 0, {
                                            'w:rPr': [
                                                { 'w:sz': '', attr: { 'w:val': fontSize } }
                                            ]
                                        } // 'w:rPr'
                                        );
                                    }
                                    else if (checkR > 1) {
                                        this.globalTable['w:tbl'][i]['w:tr'][j]['w:tc'][1]['w:p'][1]['w:r'][0]['w:rPr'].push({ 'w:sz': '', attr: { 'w:val': fontSize } });
                                    }
                                }
                            } //  if check fontSize
                            if (backgroundCell != undefined) { // if check background cell table
                                this.globalTable['w:tbl'][i]['w:tr'][j]['w:tc'][0]['w:tcPr'].push({ 'w:shd': '', attr: { 'w:val': 'clear', 'w:color': 'auto', 'w:fill': backgroundCell } });
                            } // if check background cell table
                            if (bold != undefined) { // if for check bold
                                var checkP = this.globalTable['w:tbl'][i]['w:tr'][j]['w:tc'][1]['w:p'].length;
                                if (checkP == 1) {
                                    var checkR = this.globalTable['w:tbl'][i]['w:tr'][j]['w:tc'][1]['w:p'][0]['w:r'].length;
                                    if (checkR == 1) {
                                        this.globalTable['w:tbl'][i]['w:tr'][j]['w:tc'][1]['w:p'][0]['w:r'].splice(0, 0, {
                                            'w:rPr': [
                                                { 'w:bCs': '' } // <w:bCs/>
                                            ]
                                        } // 'w:rPr'
                                        );
                                    }
                                    else if (checkR > 1) {
                                        this.globalTable['w:tbl'][i]['w:tr'][j]['w:tc'][1]['w:p'][0]['w:r'].push({ 'w:bCs': '' } // <w:bCs/>
                                        );
                                    }
                                }
                                else if (checkP > 1) {
                                    var checkR = this.globalTable['w:tbl'][i]['w:tr'][j]['w:tc'][1]['w:p'][1]['w:r'].length;
                                    if (checkR == 1) {
                                        this.globalTable['w:tbl'][i]['w:tr'][j]['w:tc'][1]['w:p'][1]['w:r'].splice(0, 0, {
                                            'w:rPr': [
                                                { 'w:bCs': '' } // <w:bCs/>
                                            ]
                                        } // 'w:rPr'
                                        );
                                        this.globalTable;
                                    }
                                    else if (checkR > 1) {
                                        this.globalTable['w:tbl'][i]['w:tr'][j]['w:tc'][1]['w:p'][1]['w:r'].push({ 'w:bCs': '' } // <w:bCs/>
                                        );
                                        this.globalTable;
                                    }
                                }
                            } // if for check bold
                        }
                    } // for j
                } // for i
                return this.globalTable;
            }
            else {
                throw "Type of style is not object";
            } // else
        }
        else {
            return this.globalTable;
        } //  else
    };
    table.prototype.callingMethod = function (body, data, style) {
        var objTr = this.createTr(body, data);
        var objTc = this.createTc(objTr, data);
        var objP = this.createPtable(objTc, data);
        var objMerge = this.createMerge(objP, data);
        var objStyle = this.tableStyle();
        return objStyle;
        // next(json2xml(objMerge,{ attributes_key:'attr' } ));
    };
    return table;
}()); //  class
exports.table = table;
//# sourceMappingURL=createTable.js.map