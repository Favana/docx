"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
var _ = require('lodash');
describe('testAll', function () {
    it('objectTest', function (done) {
        var styleOrginal = {
            fontFamily: 'B Nazanin',
            fontSize: '11',
            color: 'red',
            bold: false,
            direction: 'rtl',
            align: 'right',
            background: 'blue'
        };
        var user = {
            fontFamily: 'B Elham',
            fontColor: 'red'
        };
        var arrKeyStyle = Object.keys(styleOrginal);
        for (var i = 0; i < arrKeyStyle.length; i++) {
            if (user[arrKeyStyle[i]] != undefined) {
                styleOrginal[arrKeyStyle[i]] = user[arrKeyStyle[i]];
                styleOrginal;
            }
        }
        console.log(styleOrginal);
        done();
    }); // it
}); //  describe
//# sourceMappingURL=testObject.js.map