import {} from 'mocha';
import {fail} from 'assert';
import {Docx} from '../classes/Docx';
import Doc = Mocha.reporters.Doc;


describe('docx',  function(){

    it('fontSize_StyleP_User', function(done){
        try{
            let docx = new Docx('test_fontSize_styleP_User.docx', 'outputUnitTest/' );
            docx.createP();
            docx.addContentP('میلاد'   , {fontSize:30});
            let out = docx.generate();
            if(out == false){
                fail("Don't create File of docx");
            }else{
                console.log("Create File of Docx");
                done();
            }
        }catch(e){
            fail(e);
            done();
        }
    }); // it fontSize_StyleP_User

    it('fontSize_StyleP_default', function(done){
        try{
            let docx= new Docx('test_fontSize_styleP_default.docx', 'outputUnitTest/' );
            docx.createP();
            docx.addContentP('میلاد');
            let out = docx.generate();
            if(out == false){
                fail("Don't create File of docx");
            }else{
                console.log("Create File of Docx");
                done();
            }
        }catch(e){
            fail(e);
            done();
        }

    });




});