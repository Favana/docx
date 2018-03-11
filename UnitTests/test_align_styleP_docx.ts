import {} from 'mocha';
import {fail} from 'assert';
import {Docx} from '../classes/Docx';

describe('docx', function(){
    it('test_align_styleP_default(right)' , function(done){
        try{
            let docx = new Docx('test_align_styleP_defualt.docx', 'outputUnitTest/');
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

    });// it test_align_styleP_default(right)


    it('test_align_styleP_user_leftAlign', function(done){
        try{
            let docx = new Docx('test_align_styleP_user_left.docx', 'outputUnitTest/');
            docx.createP();
            docx.addContentP('میلاد'   ,{align:'left'});
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

    }); // it test_align_styleP_user_leftAlign


    it('test_align_styleP_user_centerAlign', function(done){
        try{
            let docx = new Docx('test_align_styleP_user_center.docx', 'outputUnitTest/');
            docx.createP();
            docx.addContentP('میلاد'   ,{align:'center'});
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

    }); // it test_align_styleP_user_centerAlign
});