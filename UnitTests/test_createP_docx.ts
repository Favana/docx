import {} from 'mocha';
import {} from '../classes/createTable';
import {Docx} from "../classes/Docx";
import {fail} from 'assert';

describe('docx', function(){

    it('CreateP', function(done){
        try{
            let docx  = new Docx('test_createP.docx','outputUnitTest/');
            docx.createP();
            let out  = docx.generate();
            if(out == false){
                fail ("Don't create File of docx");
            }else{
                console.log("Create File of Docx");
                done();
            }

        }catch(e){
            fail(e);
            done();
        }


    });// it CreateP


    it('addContentPError', function(done){
        try{
            let docx = new Docx('test_createP.docx', 'outputUnitTest/');
            docx.addContentP('میلاد');
            let out = docx.generate();
            if(out == false){
                fail("Don't create File of docx");

            }else{
                console.log("Create File of Docx");
                done();
            }

        }catch(e){
            if(e == 'createP function is undefined'){
                console.log(e);
                done();

            }else{
                fail(e);
                done();
            }
        }

    });// it addContentPError


    it('addContentP', function(done){
        try{
            let docx= new Docx('test_createP.docx', 'outputUnitTest/');
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

    });// it addContentP

    it('font_StyleP_Error For mistake send first parameter(text)', function(done){
        try{
            let docx= new Docx('test_createP.docx', 'outputUnitTest/');
            docx.createP();
            docx.addContentP({fontFamily:'B Nazanin'});
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

    });



});