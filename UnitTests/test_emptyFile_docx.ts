import {} from 'mocha';
import {docx} from '../classes/Docx';
import {fail} from 'assert';

describe('docx', function(){

    it('generate Docx File', function(done){
        try{
            let officegen = new docx('test_emptyFile.docx','outputUnitTest/');
            let out =  officegen.generate();
            if(out == false){
                fail("Don't create File of docx");

            }else{
                console.log('Create File of Docx');
                done();
            }

        }catch (e){
            fail(e);
            done();
        }

    });// it generate Docx File

    it('generate Docx File Error for null input', function(done){
        try{
            let officegen = new docx('',''); //fileName is null OR filePath is null
            let out = officegen.generate();
            if(out == false){
                fail("Don't create File of docx");
            }else{
                console.log('Create File of Docx');
                done();
            }
        }catch (e){
            if(e == "Don't Null inputs docx function OR Don't send parameters for docx function"){
                console.log(e);
                done();
            }else{
                fail(e);
            }
        }


    });// it generate Docx File Error for null input


    it('generate Docx File Error for the mistake input', function(done){
        try{
            let officegen = new docx('outputUnitTest/','test_emptyFile.docx');
            let out = officegen.generate();
            if(out == false){
                fail("Don't create File of docx");

            }else{
                console.log('Create File of Docx');
                done();
            }
        }catch (e){
            if(e == "Don't Null inputs docx function OR Don't send parameters for docx function"){
                console.log(e);
                done();
            }else{
                fail(e);
            }
        }
    });// it generate Docx File Error for the mistake input

    it('generate Docx File Error for the filePath is null input', function(done){
        try{
            let officegen = new docx('test_emptyFile.docx','');
            let out =  officegen.generate();
            if(out == false){
                fail("Don't create File of docx");

            }else{
                console.log('Create File of Docx');
                done();
            }

        }catch (e){
            if(e == "Don't Null inputs docx function OR Don't send parameters for docx function"){
                console.log(e);
                done();

            }else{
                fail(e);
                done();
            }

        }
    });// it generate Docx File Error for the filePath is null input

    it('generate Docx File Error for the fileName is null input', function(done){
        try{
            let officegen = new docx('','outputUnitTest/');
            let out =  officegen.generate();
            if(out == false){
                fail("Don't create File of docx");

            }else{
                console.log('Create File of Docx');
                done();
            }
        }catch (e){
            if(e == "Don't Null inputs docx function OR Don't send parameters for docx function"){
                console.log(e);
                done();

            }else{
                fail(e);
                done();
            }

        }
    });



}); // describe