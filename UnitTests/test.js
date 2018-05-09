"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
var _ = require('lodash');
describe('testData', function () {
    it('test', function (done) {
        var data = [
            { x: 1, y: 0, value: '', mergeRow: '', mergeCol: '', style: { fontFamily: 'B Nazanin', fontSize: 18, fontColor: 'white', align: 'center', background: 'black', borderSize: 30 } },
            { x: 1, y: 1, value: 'سال 1390', mergeRow: '', mergeCol: '' },
            { x: 1, y: 2, value: 'سال1391', mergeRow: '', mergeCol: '' },
            { x: 2, y: 0, value: 'کل', mergeRow: '', mergeCol: '' },
            { x: 2, y: 1, value: '21545288', mergeRow: '', mergeCol: '', style: { fontFamily: 'B Nazanin', fontSize: 18, fontColor: 'white', align: 'center', background: 'black', borderSize: 30 } },
            { x: 2, y: 2, value: '85487525', mergeRow: '', mergeCol: '' },
            { x: 3, y: 0, value: 'البرز', mergeRow: '', mergeCol: '' },
            { x: 3, y: 1, value: '2521', mergeRow: '', mergeCol: '', style: { fontFamily: 'B Nazanin', fontSize: 18, fontColor: 'white', align: 'center', background: 'black', borderSize: 30 } },
            { x: 3, y: 2, value: '5485', mergeRow: '', mergeCol: '' },
            { x: 4, y: 0, value: 'بندرعباس', mergeRow: '', mergeCol: '', style: { fontFamily: 'B Nazanin' } },
            { x: 4, y: 1, value: '145214', mergeRow: '', mergeCol: '' },
            { x: 4, y: 2, value: '2255', mergeRow: '', mergeCol: '' }
        ]; // data
        var count = data.length;
        // let styleData = {};
        var styleObject = {};
        var styleData = [];
        for (var i = 0; i <= count; i++) {
            var infoData = _.find(data, data[i]);
            var x = infoData.x;
            var y = infoData.y;
            var style = infoData.style;
            x;
            y;
            style;
            var check = Object.keys(styleObject).length;
            check;
            if (check == 0) {
                styleObject = { x: x, y: y, style: style };
                styleData.push(styleObject);
            }
            else if (check > 0) {
                //  this.globalP = (<any>Object).assign(objP, this.globalP);
                styleObject = { x: x, y: y, style: style };
                styleData.push(styleObject);
            }
            //let styleData = {x:x, y:y, style:style};
            styleData;
        } //  for
        done();
    }); //  it
}); // describe
//# sourceMappingURL=test.js.map