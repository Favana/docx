import {} from 'mocha';
import {fail} from 'assert';
import {docx} from '../classes/Docx';


describe('docx', function(){

    it('font_StyleP_User', function(done){
        let officegen = new docx('test_font_styleP_User.docx', 'outputUnitTest/');
        officegen.createP();
        officegen.addContentP('میلاد'  , {fontFamily:'B Elham'});
        let out = officegen.generate();
        if(out == false){
            fail("Don't create File of docx");
        }else{
            console.log("Create File of Docx");
            done();
        }
    });// it for font_StyleP_User

    it('font_StyleP_default', function(done){
        let officegen = new docx('test_font_styleP.docx', 'outputUnitTest/');
        officegen.createP();
        officegen.addContentP('میلاد');
        let out = officegen.generate();
        if(out == false){
            fail("Don't create File of docx");
        }else{
            console.log("Create File of Docx");
            done();
        }
    });// it for font_StyleP_default
});