import {} from 'mocha';
import {docx} from '../classes/Docx';
import {fail} from 'assert';

describe('docx', function(){

    it('Create File Docx', function(done){

        let docxOffice = new docx('test_emptyFile.docx','outputUnitTest/');
        let out =  docxOffice.generate();
        if(out == false){
            fail("Don't create File of docx");

        }else{
            console.log('Create File of Docx');
            done();
        }


    })
}); // describe