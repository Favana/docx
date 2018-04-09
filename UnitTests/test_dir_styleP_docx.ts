import {} from 'mocha';
import {fail} from 'assert';
import {Docx} from '../classes/Docx';

describe('docx', function(){

    it('direction_styleP_default', function(done){
        try{
            let docx = new Docx('test_dir_styleP_default.docx', 'outputUnitTest/');
            docx.createP();
            docx.addContentP('میلاد' , { fontSize:30});
            docx.addContentP('میلاد');

            let out = docx.generate();
            if(out == false){
                fail("Don't create File of docx");
            }else{
                console.log("Create File of Docx");
                done();
            }
        }catch (e){

        }
    });//  it direction_styleP_default

    it('direction_styleP_User', function(done){
        try{
            let docx = new Docx('test_dir_styleP_user.docx', 'outputUnitTest/');
            docx.createP();
            docx.addContentP( 'milad' , {direction:'ltr', fontColor:'red', fontSize:30, fontFamily:'Times New Roman'});
            docx.addContentP( 'milad' , {direction:'ltr'});
            let out = docx.generate();
            if(out == false){
                fail("Don't create File of docx");
            }else{
                console.log("Create File of Docx");
                done();
            }

        }catch (e){
            fail(e);
            done();
        }

    });// it direction_styleP_user


});