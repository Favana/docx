import {} from 'mocha';
import {} from '../classes/createTable';
import {docx} from "../classes/Docx";
import {fail} from 'assert';

describe('docx', function(){

    it('CreateP', function(done){
        let officegen  = new docx('test_createP.docx','outputUnitTest/');
        officegen.createP();
        let out  = officegen.generate();
        if(out == false){
            fail ("Don't create File of docx");
        }else{
            console.log("Create File of Docx");
            done();
        }

    });// it CreateP


    it('addContentPError', function(done){
        let officegen = new docx('test_createP.docx', 'outputUnitTest/');
        officegen.addContentP('میلاد');
        let out = officegen.generate();
        if(out == false){
            fail("Don't create File of docx");

        }else{
            console.log("Create File of Docx");
            done();
        }

    });// it addContentPError


    it('addContentP', function(done){
        let officegen = new docx('test_createP.docx', 'outputUnitTest/');
        officegen.createP();
        officegen.addContentP('میلاد');
        let out = officegen.generate();
        if(out == false){
            fail("Don't create File of docx");
        }else{
            console.log("Create File of Docx");
            done();
        }


    });// it addContentP


});