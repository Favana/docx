import {} from 'mocha';
import {fail} from 'assert';
import {Docx} from '../classes/Docx';
let _ = require('lodash');

describe('style in table with new data', function(){


    it('style Without merge', function(done){
        let data = [
            {x: 1, y: 0, value: '',mergeRow:'', mergeCol:'' , style:{borderSize: 30}},   //mergeRow:(x) ,,,, mergeCol:(y)

            {x: 1, y: 1, value: 'سال 1390', mergeRow:'', mergeCol:''},

            {x: 1, y: 2, value: 'سال1391',mergeRow:'', mergeCol:''},

            {x: 2, y: 0, value: 'کل',mergeRow:'', mergeCol:'', style:{borderSize: 30}},

            {x: 2, y: 1, value: '21545288', mergeRow:'', mergeCol:''},

            {x: 2, y: 2, value: '85487525',mergeRow:'', mergeCol:''},

            {x: 3, y: 0, value: 'البرز',mergeRow:'', mergeCol:'',style:{fontFamily: 'B Nazanin', align:'right'}},

            {x: 3, y: 1, value: '2521',mergeRow:'', mergeCol:''},

            {x: 3, y: 2, value: '5485',mergeRow:'', mergeCol:''},

            {x: 4, y: 0, value: 'بندرعباس',mergeRow:'', mergeCol:''},

            {x: 4, y: 1, value: '145214',mergeRow:'', mergeCol:''},

            {x: 4, y: 2, value: '2255',mergeRow:'', mergeCol:''}

        ];// data

        let docx = new Docx('test_newData_without_style_tbl.docx', 'outputUnitTest/');
        docx.createTable(data);
        let out = docx.generate();
        if(out == false){
            fail("Don't create File of docx");
        }else{
            console.log("Create File of Docx");
            done();
        }



    });//  it for style Without merge

});// describe