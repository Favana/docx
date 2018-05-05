import {} from 'mocha';
import {fail} from 'assert';
import {Docx} from '../classes/Docx';


describe('style table', function(){

    it('test_font_style_tbl_docx', function(done){
        try{
            let data = [
                {x: 1, y: 0, value: '',mergeRow:'', mergeCol:''},       //mergeRow:(x) ,,,, mergeCol:(y)
                {x: 1, y: 1, value: 'سال 1390', mergeRow:'', mergeCol:''},
                {x: 1, y: 2, value: 'سال1391',mergeRow:'', mergeCol:''},

                {x: 2, y: 0, value: 'کل',mergeRow:'', mergeCol:''},
                {x: 2, y: 1, value: '21545288', mergeRow:'', mergeCol:''},
                {x: 2, y: 2, value: '85487525',mergeRow:'', mergeCol:''},

                {x: 3, y: 0, value: 'البرز',mergeRow:'', mergeCol:''},
                {x: 3, y: 1, value: '2521',mergeRow:'', mergeCol:''},
                {x: 3, y: 2, value: '5485',mergeRow:'', mergeCol:''},

                {x: 4, y: 0, value: 'بندرعباس',mergeRow:'', mergeCol:''},
                {x: 4, y: 1, value: '145214',mergeRow:'', mergeCol:''},
                {x: 4, y: 2, value: '2255',mergeRow:'', mergeCol:''}

            ];// data

            let style = [
                {x: 1, y: 0},
                {x: 1, y: 1},
                {x: 1, y: 2},
                {x: 2, y: 0, fontFamily :'B Nazanin'},
                {x: 2, y: 1},
                {x: 2, y: 2},
                {x: 3, y: 0, fontFamily :'B Nazanin'},
                {x: 3, y: 1},
                {x: 3, y: 2},
                {x: 4, y: 0, fontFamily :'B Nazanin'},
                {x: 4, y: 1},
                {x: 4, y: 2}
            ];// style
            let docx = new Docx('test_font_style_tbl.docx', 'outputUnitTest/');
            docx.createTable(data, style);
            let out = docx.generate();
            if(out == false){
                fail("Don't create File of docx");
            }else{
                console.log("Create File of Docx");
                done();
            }

        }catch (e) {
            fail(e);
            done();
        }

    });//  it test_font_style_docx

});// describe