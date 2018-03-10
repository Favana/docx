import {} from 'mocha';
import {fail} from 'assert';
import {docx} from '../classes/Docx';


describe('docx',  function(){

    it('fontSize_StyleP_User', function(done){
        let officegen = new docx('test_fontSize_styleP_User.docx', 'outputUnitTest/' );
        officegen.createP();
        officegen.addContentP('میلاد'   , {fontSize:30});
        let out = officegen.generate();
        if(out == false){
            fail("Don't create File of docx");
        }else{
            console.log("Create File of Docx");
            done();
        }

    }); // it fontSize_StyleP_User

    it('fontSize_StyleP_default', function(done){
        let officegen = new docx('test_fontSize_styleP_default.docx', 'outputUnitTest/' );
        officegen.createP();
        officegen.addContentP('میلاد');
        let out = officegen.generate();
        if(out == false){
            fail("Don't create File of docx");
        }else{
            console.log("Create File of Docx");
            done();
        }


    });



});