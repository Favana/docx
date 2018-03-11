import {} from 'mocha';
import {fail} from 'assert';
import {Docx} from '../classes/Docx';

describe('docx', function(){

    it('bold_styleP_default', function(done){
        try{
            let docx = new Docx('test_bold_styleP_default.docx', 'outputUnitTest/');
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

    }); // it test_bold_styleP_docx


    it('bold_styleP_user', function(done){
        try{
            let docx = new Docx('test_bold_styleP_user.docx', 'outputUnitTest/');
            docx.createP();
            docx.addContentP('میلاد' ,  {bold: 'true'});
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