import {} from 'mocha';
import {isUndefined} from "util";
let _ = require('lodash');

describe('testAll', function(){

    it('objectTest', function(done){
        let  styleOrginal= {
            fontFamily:'B Nazanin',
            fontSize : '11',
            color:'red',
            bold: false,
            direction: 'rtl',
            align : 'right',
            background: 'blue'
        };

     let user = {
        fontFamily : 'B Elham',
         fontColor: 'red'
     };

     let arrKeyStyle = Object.keys(styleOrginal);


       for(let i=0; i<arrKeyStyle.length; i++){
           if(user[arrKeyStyle[i]] != undefined){
               styleOrginal[arrKeyStyle[i]] = user[arrKeyStyle[i]];
               styleOrginal;
           }

       }



    console.log(styleOrginal);



     done()
    });// it

});//  describe