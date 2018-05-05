import {} from 'mocha';
import {fail} from 'assert';
import {Docx} from '../classes/Docx';



describe('style table', function(){
    it('test_textAlign_Center_style_tbl', function(done){
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
                {x: 1, y: 1, align: 'center'},
                {x: 1, y: 2,  align:'center'},
                {x: 2, y: 0,  align:'center'},
                {x: 2, y: 1,  align:'center'},
                {x: 2, y: 2,  align:'center'},
                {x: 3, y: 0, align:'center'},
                {x: 3, y: 1,  align:'center'},
                {x: 3, y: 2 ,  align:'center'},
                {x: 4, y: 0, align:'center'},
                {x: 4, y: 1,  align:'center'},
                {x: 4, y: 2,  align:'center'}
            ];// style

            let docx = new Docx('test_textAlign_Center_style_tbl.docx', 'outputUnitTest/');
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

    });// it test_textAlign_Center_style_tbl


    it('test_textAlign_Right_style_tbl', function(done){

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
                {x: 1, y: 1, align: 'right'},
                {x: 1, y: 2,  align:'right'},
                {x: 2, y: 0,  align:'right'},
                {x: 2, y: 1,  align:'right'},
                {x: 2, y: 2,  align:'right'},
                {x: 3, y: 0, align:'right'},
                {x: 3, y: 1,  align:'right'},
                {x: 3, y: 2 ,  align:'right'},
                {x: 4, y: 0, align:'right'},
                {x: 4, y: 1,  align:'right'},
                {x: 4, y: 2,  align:'right'}
            ];// style


            let docx = new Docx('test_textAlign_Right_style_tbl.docx', 'outputUnitTest/');
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

    });//  it test_textAlign_Right_style_tbl
});// describe