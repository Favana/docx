import {} from 'mocha';
import {fail} from 'assert';
import {Docx} from '../classes/Docx';


describe('docx', function(){

    it('font_StyleP_User', function(done){
        try{
            let docx = new Docx('test_font_styleP_User.docx', 'outputUnitTest/');
            docx.createP();
            docx.addContentP('میلاد'  , {fontFamily:'B Elham'});
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

    });// it for font_StyleP_User

    it('font_StyleP_default', function(done){
        try{
            let docx = new Docx('test_font_styleP_default.docx', 'outputUnitTest/');
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

    });// it for font_StyleP_default

    it('font_StyleP_Error For mistake send first parameter(text)', function(done){
        try{
            let docx =  new Docx('test_font_styleP_User.docx', 'outputUnitTest/');
            docx.createP();
            docx.addContentP({fontFamily: 'B Elham'});
            let out = docx.generate();
            if(out == false){
                fail("Don't create File of docx");
            }else{
                console.log("Create File of Docx");
                done();
            }
        }catch(e){
            if(e == 'The sending parameter is incorrect'){
                console.log(e);
                done();
            }else if(e == 'createP function is undefined'){
                console.log(e);
                done();
            }else{
                fail(e);
                done();
            }
        }

    });//  it font_StyleP_Error For mistake send first parameter


    it('font_StyleP_Error For mistake send second parameter(object)', function(done){
        try{
            let docx =  new Docx('test_font_styleP_User.docx', 'outputUnitTest/');
            docx.createP();
            docx.addContentP('میلاد', 'error');
            let out = docx.generate();
            if(out == false){
                fail("Don't create File of docx");
            }else{
                console.log("Create File of Docx");
                done();
            }
        }catch(e){
            if(e == 'The sending parameter is incorrect' ){
                console.log(e);
                done();
            }else if( e == 'createP function is undefined'){
                console.log(e);
                done();
            }else{
                fail(e);
                done();
            }
        }
    });// it font_StyleP_Error For mistake send second parameter












});