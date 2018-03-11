import {} from 'mocha';
import {fail} from 'assert';
import {Docx} from '../classes/Docx';

describe('docx', function(){

    it('background_styleP_default', function(done){
        try{
            let docx = new Docx('test_background_styleP_default.docx', 'outputUnitTest/');
            docx.createP();
            docx.addContentP('میلاد');
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

    });// it  background_styleP_default

    it('background_styleP_user', function(done){
        try{
            let docx = new Docx('test_background_styleP_user.docx', 'outputUnitTest/');
            docx.createP();
            docx.addContentP('میلاد' , {fontColor:'white', backgroundFont:'black'});
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

    });


});