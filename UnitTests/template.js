"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
var assert_1 = require("assert");
var Docx_1 = require("../classes/Docx");
var _ = require('lodash');
describe('default Template', function () {
    /* it('Template One' ,function(done){ /// فایل مربوط به سازمان جغرافیایی نیرو های مسلح
       let data = [
           {x: 1, y: 0, value: 'استان',mergeRow:'3', mergeCol:'',style:{background:'#DCDCDC',align:'center',fontFamily:'B Nazanin'}},   //mergeRow:(x) ,,,, mergeCol:(y)
           {x: 1, y: 1, value: 'طول شرقی', mergeRow:'', mergeCol:'4',  style:{background:'#DCDCDC',align:'center',fontFamily:'B Nazanin'}},
           {x: 1, y: 2, value: '',mergeRow:'', mergeCol:''},
           {x: 1, y: 3, value: '',mergeRow:'', mergeCol:''},
           {x: 1, y: 4, value: '',mergeRow:'', mergeCol:''},
 
 
           {x: 2, y: 0, value: '', mergeRow:'', mergeCol:''},// merge
           {x: 2, y: 1, value: 'حداقل',mergeRow:'', mergeCol:'2',style:{background:'#DCDCDC',align:'center',fontFamily:'B Nazanin'}},
           {x: 2, y: 2, value: '',mergeRow:'', mergeCol:''},
           {x: 2, y: 3, value: 'حداکثر',mergeRow:'', mergeCol:'2',style:{background:'#DCDCDC',align:'center',fontFamily:'B Nazanin'}},
           {x: 2, y: 4, value: '',mergeRow:'', mergeCol:''},
 
 
           {x: 3, y: 0, value: '',mergeRow:'', mergeCol:''},// merge
           {x: 3, y: 1, value: 'دقیقه',mergeRow:'', mergeCol:'',style:{background:'#DCDCDC',align:'center',fontFamily:'B Nazanin'}},
           {x: 3, y: 2, value: 'درجه',mergeRow:'', mergeCol:'',style:{background:'#DCDCDC',align:'center',fontFamily:'B Nazanin'}},
           {x: 3, y: 3, value: 'دقیقه',mergeRow:'', mergeCol:'',style:{background:'#DCDCDC',align:'center',fontFamily:'B Nazanin'}},
           {x: 3, y: 4, value: 'درجه',mergeRow:'', mergeCol:'',style:{background:'#DCDCDC',align:'center',fontFamily:'B Nazanin'}},
 
 
           {x: 4, y: 0, value: 'کل کشور',mergeRow:'', mergeCol:'',style:{background:'#DCDCDC',align:'center',fontFamily:'B Nazanin'}},
           {x: 4, y: 1, value: '4',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
           {x: 4, y: 2, value: '2',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
           {x: 4, y: 3, value: '145214',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
           {x: 4, y: 4, value: '145214',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
 
 
           {x: 5, y: 0, value: 'تهران',mergeRow:'', mergeCol:'',style:{background:'#DCDCDC',align:'center',fontFamily:'B Nazanin'}},
           {x: 5, y: 1, value: '145214',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
           {x: 5, y: 2, value: '145214',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
           {x: 5, y: 3, value: '145214',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
           {x: 5, y: 4, value: '145214',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
 
 
           {x: 6, y: 0, value: 'گیلان',mergeRow:'', mergeCol:'',style:{background:'#DCDCDC',align:'center',fontFamily:'B Nazanin'}},
           {x: 6, y: 1, value: '145214',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
           {x: 6, y: 2, value: '145214',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
           {x: 6, y: 3, value: '145214',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
           {x: 6, y: 4, value: '145214',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
 
 
           {x: 7, y: 0, value: 'مازندران',mergeRow:'', mergeCol:'',style:{background:'#DCDCDC',align:'center',fontFamily:'B Nazanin'}},
           {x: 7, y: 1, value: '145214',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
           {x: 7, y: 2, value: '145214',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
           {x: 7, y: 3, value: '145214',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
           {x: 7, y: 4, value: '145214',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
 
 
 
           {x: 8, y: 0, value: 'البرز',mergeRow:'', mergeCol:'',style:{background:'#DCDCDC',align:'center',fontFamily:'B Nazanin'}},
           {x: 8, y: 1, value: '145214',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
           {x: 8, y: 2, value: '145214',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
           {x: 8, y: 3, value: '145214',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
           {x: 8, y: 4, value: '145214',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}}
 
 
       ];// data / داده مربوط به سازمان جغرافیایی نیرو های مسلح //
 
       let docx = new Docx('temp1.docx', 'outputUnitTestTmp/');
       docx.createTable(data);
       let out = docx.generate();
       if(out == false){
           fail("Don't create File of docx");
       }else{
           console.log("Create File of Docx");
           done();
       }
 
   });//  it   /// فایل مربوط به سازمان جغرافیایی نیرو های مسلح*/
    /*it('Template Two', function(done){ ///فایل مربوط به مراكز ارائه دهنده مراقبت هاي اوليه بهداشتي ///

         let data = [
             {x: 1, y: 0, value: 'سال و استان',mergeRow:'2', mergeCol:'2',style:{background:'#DCDCDC',align:'center',fontFamily:'B Nazanin'}},   //mergeRow:(x) ,,,, mergeCol:(y)
             {x: 1, y: 1, value: '',mergeRow:'', mergeCol:''},
             {x: 1, y: 2, value: 'مراکز بهداشتی و درمانی',mergeRow:'', mergeCol:'3',style:{background:'#DCDCDC',align:'center',fontFamily:'B Nazanin'}},
             {x: 1, y: 3, value: '', mergeRow:'', mergeCol:''},
             {x: 1, y: 4, value: '',mergeRow:'', mergeCol:''},
             {x: 1, y: 5, value: 'پایگاه بهداشت',mergeRow:'', mergeCol:'3',style:{background:'#DCDCDC',align:'center',fontFamily:'B Nazanin'}},
             {x: 1, y: 6, value: '',mergeRow:'', mergeCol:''},
             {x: 1, y: 7, value: '',mergeRow:'', mergeCol:''},
             {x: 1, y: 8, value: 'تسهیلات زایمانی',mergeRow:'2', mergeCol:'2',style:{background:'#DCDCDC',align:'center',fontFamily:'B Nazanin'}},
             {x: 1, y: 9, value: '',mergeRow:'', mergeCol:''},
             {x: 1, y: 10, value: 'خانه هاي بهداشت فعال',mergeRow:'2', mergeCol:'3',style:{background:'#DCDCDC',align:'center',fontFamily:'B Nazanin'}},
             {x: 1, y: 11, value: '',mergeRow:'', mergeCol:''},
             {x: 1, y: 12, value: '',mergeRow:'', mergeCol:''},






             {x: 2, y: 0, value: '',mergeRow:'', mergeCol:''},   //mergeRow:(x) ,,,, mergeCol:(y)
             {x: 2, y: 1, value: '',mergeRow:'', mergeCol:''},
             {x: 2, y: 2, value: 'جمع',mergeRow:'', mergeCol:'',style:{background:'#DCDCDC',align:'center',fontFamily:'B Nazanin'}},
             {x: 2, y: 3, value: 'شهری', mergeRow:'', mergeCol:'',style:{background:'#DCDCDC',align:'center',fontFamily:'B Nazanin'}},
             {x: 2, y: 4, value: 'روستایی',mergeRow:'', mergeCol:'',style:{background:'#DCDCDC',align:'center',fontFamily:'B Nazanin'}},
             {x: 2, y: 5, value: 'جمع',mergeRow:'', mergeCol:'',style:{background:'#DCDCDC',align:'center',fontFamily:'B Nazanin'}},
             {x: 2, y: 6, value: 'شهری',mergeRow:'', mergeCol:'',style:{background:'#DCDCDC',align:'center',fontFamily:'B Nazanin'}},
             {x: 2, y: 7, value: 'روستایی',mergeRow:'', mergeCol:'',style:{background:'#DCDCDC',align:'center',fontFamily:'B Nazanin'}},
             {x: 2, y: 8, value: '',mergeRow:'', mergeCol:''},
             {x: 2, y: 9, value: '',mergeRow:'', mergeCol:''},
             {x: 2, y: 10, value: '',mergeRow:'', mergeCol:''},
             {x: 2, y: 11, value: '',mergeRow:'', mergeCol:''},
             {x: 2, y: 12, value: '',mergeRow:'', mergeCol:''},




             {x: 3, y: 0, value: '1375',mergeRow:'', mergeCol:'2',style:{background:'#DCDCDC',align:'center',fontFamily:'B Nazanin'}},// merge
             {x: 3, y: 1, value: '',mergeRow:'', mergeCol:''},
             {x: 3, y: 2, value: '...',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 3, y: 3, value: '...',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 3, y: 4, value: '...',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 3, y: 5, value: '...',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 3, y: 6, value: '...',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 3, y: 7, value: '...',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},//
             {x: 3, y: 8, value: '...',mergeRow:'', mergeCol:'2',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 3, y: 9, value: '',mergeRow:'', mergeCol:''},
             {x: 3, y: 10, value: '13314',mergeRow:'', mergeCol:'3',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 3, y: 11, value: '',mergeRow:'', mergeCol:''},
             {x: 3, y: 12, value: '',mergeRow:'', mergeCol:''},



             {x: 4, y: 0, value: '1380',mergeRow:'', mergeCol:'2',style:{background:'#DCDCDC',align:'center',fontFamily:'B Nazanin'}},// merge
             {x: 4, y: 1, value: '',mergeRow:'', mergeCol:''},
             {x: 4, y: 2, value: '...',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 4, y: 3, value: '...',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 4, y: 4, value: '...',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 4, y: 5, value: '...',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 4, y: 6, value: '...',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 4, y: 7, value: '...',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},//
             {x: 4, y: 8, value: '...',mergeRow:'', mergeCol:'2',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 4, y: 9, value: '',mergeRow:'', mergeCol:''},
             {x: 4, y: 10, value: '...',mergeRow:'', mergeCol:'3',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 4, y: 11, value: '',mergeRow:'', mergeCol:''},
             {x: 4, y: 12, value: '',mergeRow:'', mergeCol:''},



             {x: 5, y: 0, value: '1385',mergeRow:'', mergeCol:'2',style:{background:'#DCDCDC',align:'center',fontFamily:'B Nazanin'}},// merge
             {x: 5, y: 1, value: '',mergeRow:'', mergeCol:''},
             {x: 5, y: 2, value: '...',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 5, y: 3, value: '...',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 5, y: 4, value: '...',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 5, y: 5, value: '...',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 5, y: 6, value: '...',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 5, y: 7, value: '...',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},//
             {x: 5, y: 8, value: '...',mergeRow:'', mergeCol:'2',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 5, y: 9, value: '',mergeRow:'', mergeCol:''},
             {x: 5, y: 10, value: '12166',mergeRow:'', mergeCol:'3',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 5, y: 11, value: '',mergeRow:'', mergeCol:''},
             {x: 5, y: 12, value: '',mergeRow:'', mergeCol:''},



             {x: 6, y: 0, value: '1392',mergeRow:'', mergeCol:'2',style:{background:'#DCDCDC',align:'center',fontFamily:'B Nazanin'}},// merge
             {x: 6, y: 1, value: '',mergeRow:'', mergeCol:''},
             {x: 6, y: 2, value: '...',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 6, y: 3, value: '...',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 6, y: 4, value: '...',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 6, y: 5, value: '...',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 6, y: 6, value: '...',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 6, y: 7, value: '...',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},//
             {x: 6, y: 8, value: '...',mergeRow:'', mergeCol:'2',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 6, y: 9, value: '',mergeRow:'', mergeCol:''},
             {x: 6, y: 10, value: '17749',mergeRow:'', mergeCol:'3',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 6, y: 11, value: '',mergeRow:'', mergeCol:''},
             {x: 6, y: 12, value: '',mergeRow:'', mergeCol:''},



             {x: 7, y: 0, value: '1393',mergeRow:'', mergeCol:'2',style:{background:'#DCDCDC',align:'center',fontFamily:'B Nazanin'}},// merge
             {x: 7, y: 1, value: '',mergeRow:'', mergeCol:''},
             {x: 7, y: 2, value: '4657',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 7, y: 3, value: '2376',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 7, y: 4, value: '2281',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 7, y: 5, value: '2401',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 7, y: 6, value: '2278',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 7, y: 7, value: '123',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},//
             {x: 7, y: 8, value: '274',mergeRow:'', mergeCol:'2',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 7, y: 9, value: '',mergeRow:'', mergeCol:''},
             {x: 7, y: 10, value: '17644',mergeRow:'', mergeCol:'3',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 7, y: 11, value: '',mergeRow:'', mergeCol:''},
             {x: 7, y: 12, value: '',mergeRow:'', mergeCol:''},


             {x: 8, y: 0, value: '1394',mergeRow:'', mergeCol:'2',style:{background:'#DCDCDC',align:'center',fontFamily:'B Nazanin'}},// merge
             {x: 8, y: 1, value: '',mergeRow:'', mergeCol:''},
             {x: 8, y: 2, value: '4672',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 8, y: 3, value: '2394',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 8, y: 4, value: '2278',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 8, y: 5, value: '2915',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 8, y: 6, value: '2776',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 8, y: 7, value: '139',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},//
             {x: 8, y: 8, value: '257',mergeRow:'', mergeCol:'2',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 8, y: 9, value: '',mergeRow:'', mergeCol:''},
             {x: 8, y: 10, value: '17852',mergeRow:'', mergeCol:'3',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 8, y: 11, value: '',mergeRow:'', mergeCol:''},
             {x: 8, y: 12, value: '',mergeRow:'', mergeCol:''},



             {x: 9, y: 0, value: '1395',mergeRow:'', mergeCol:'2',style:{background:'#DCDCDC',align:'center',fontFamily:'B Nazanin'}},// merge
             {x: 9, y: 1, value: '',mergeRow:'', mergeCol:''},
             {x: 9, y: 2, value: '',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 9, y: 3, value: '',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 9, y: 4, value: '',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 9, y: 5, value: '',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 9, y: 6, value: '',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 9, y: 7, value: '',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},//
             {x: 9, y: 8, value: '',mergeRow:'', mergeCol:'2',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 9, y: 9, value: '',mergeRow:'', mergeCol:''},
             {x: 9, y: 10, value: '',mergeRow:'', mergeCol:'3',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 9, y: 11, value: '',mergeRow:'', mergeCol:''},
             {x: 9, y: 12, value: '',mergeRow:'', mergeCol:''},




             {x: 10, y: 0, value: '1396',mergeRow:'', mergeCol:'2',style:{background:'#DCDCDC',align:'center',fontFamily:'B Nazanin'}},// merge
             {x: 10, y: 1, value: '',mergeRow:'', mergeCol:''},
             {x: 10, y: 2, value: '',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 10, y: 3, value: '',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 10, y: 4, value: '',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 10, y: 5, value: '',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 10, y: 6, value: '',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 10, y: 7, value: '',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},//
             {x: 10, y: 8, value: '',mergeRow:'', mergeCol:'2',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 10, y: 9, value: '',mergeRow:'', mergeCol:''},
             {x: 10, y: 10, value: '',mergeRow:'', mergeCol:'3',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 10, y: 11, value: '',mergeRow:'', mergeCol:''},
             {x: 10, y: 12, value: '',mergeRow:'', mergeCol:''},




             {x: 11, y: 0, value: 'آذربايجان شرقي',mergeRow:'', mergeCol:'2',style:{background:'#DCDCDC',align:'center',fontFamily:'B Nazanin'}},// merge
             {x: 11, y: 1, value: '',mergeRow:'', mergeCol:''},
             {x: 11, y: 2, value: '280',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 11, y: 3, value: '145',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 11, y: 4, value: '135',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 11, y: 5, value: '69',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 11, y: 6, value: '49',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 11, y: 7, value: '20',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},//
             {x: 11, y: 8, value: '14',mergeRow:'', mergeCol:'2',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 11, y: 9, value: '',mergeRow:'', mergeCol:''},
             {x: 11, y: 10, value: '1072',mergeRow:'', mergeCol:'3',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 11, y: 11, value: '',mergeRow:'', mergeCol:''},
             {x: 11, y: 12, value: '',mergeRow:'', mergeCol:''},



             {x: 12, y: 0, value: 'آذربايجان غربی',mergeRow:'', mergeCol:'2',style:{background:'#DCDCDC',align:'center',fontFamily:'B Nazanin'}},// merge
             {x: 12, y: 1, value: '',mergeRow:'', mergeCol:''},
             {x: 12, y: 2, value: '224',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 12, y: 3, value: '109',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 12, y: 4, value: '115',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 12, y: 5, value: '115',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 12, y: 6, value: '109',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 12, y: 7, value: '6',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},//
             {x: 12, y: 8, value: '12',mergeRow:'', mergeCol:'2',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 12, y: 9, value: '',mergeRow:'', mergeCol:''},
             {x: 12, y: 10, value: '985',mergeRow:'', mergeCol:'3',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 12, y: 11, value: '',mergeRow:'', mergeCol:''},
             {x: 12, y: 12, value: '',mergeRow:'', mergeCol:''},



             {x: 13, y: 0, value: 'اردبیل',mergeRow:'', mergeCol:'2',style:{background:'#DCDCDC',align:'center',fontFamily:'B Nazanin'}},// merge
             {x: 13, y: 1, value: '',mergeRow:'', mergeCol:''},
             {x: 13, y: 2, value: '102',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 13, y: 3, value: '51',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 13, y: 4, value: '51',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 13, y: 5, value: '60',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 13, y: 6, value: '60',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 13, y: 7, value: '0',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},//
             {x: 13, y: 8, value: '8',mergeRow:'', mergeCol:'2',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 13, y: 9, value: '',mergeRow:'', mergeCol:''},
             {x: 13, y: 10, value: '509',mergeRow:'', mergeCol:'3',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 13, y: 11, value: '',mergeRow:'', mergeCol:''},
             {x: 13, y: 12, value: '',mergeRow:'', mergeCol:''},



             {x: 14, y: 0, value: 'اصفهان',mergeRow:'', mergeCol:'2',style:{background:'#DCDCDC',align:'center',fontFamily:'B Nazanin'}},// merge
             {x: 14, y: 1, value: '',mergeRow:'', mergeCol:''},
             {x: 14, y: 2, value: '293',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 14, y: 3, value: '204',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 14, y: 4, value: '89',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 14, y: 5, value: '260',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 14, y: 6, value: '260',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 14, y: 7, value: '0',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},//
             {x: 14, y: 8, value: '6',mergeRow:'', mergeCol:'2',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 14, y: 9, value: '',mergeRow:'', mergeCol:''},
             {x: 14, y: 10, value: '559',mergeRow:'', mergeCol:'3',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 14, y: 11, value: '',mergeRow:'', mergeCol:''},
             {x: 14, y: 12, value: '',mergeRow:'', mergeCol:''},





             {x: 15, y: 0, value: 'البرز',mergeRow:'', mergeCol:'2',style:{background:'#DCDCDC',align:'center',fontFamily:'B Nazanin'}},// merge
             {x: 15, y: 1, value: '',mergeRow:'', mergeCol:''},
             {x: 15, y: 2, value: '68',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 15, y: 3, value: '54',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 15, y: 4, value: '14',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 15, y: 5, value: '126',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 15, y: 6, value: '125 ',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 15, y: 7, value: '1',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},//
             {x: 15, y: 8, value: '1',mergeRow:'', mergeCol:'2',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 15, y: 9, value: '',mergeRow:'', mergeCol:''},
             {x: 15, y: 10, value: '80',mergeRow:'', mergeCol:'3',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 15, y: 11, value: '',mergeRow:'', mergeCol:''},
             {x: 15, y: 12, value: '',mergeRow:'', mergeCol:''},



             {x: 16, y: 0, value: 'ایلام',mergeRow:'', mergeCol:'2',style:{background:'#DCDCDC',align:'center',fontFamily:'B Nazanin'}},// merge
             {x: 16, y: 1, value: '',mergeRow:'', mergeCol:''},
             {x: 16, y: 2, value: '66',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 16, y: 3, value: '39',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 16, y: 4, value: '27',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 16, y: 5, value: '50',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 16, y: 6, value: '50 ',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 16, y: 7, value: '0',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},//
             {x: 16, y: 8, value: '7',mergeRow:'', mergeCol:'2',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 16, y: 9, value: '',mergeRow:'', mergeCol:''},
             {x: 16, y: 10, value: '204',mergeRow:'', mergeCol:'3',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 16, y: 11, value: '',mergeRow:'', mergeCol:''},
             {x: 16, y: 12, value: '',mergeRow:'', mergeCol:''},



             {x: 17, y: 0, value: 'بوشهر',mergeRow:'', mergeCol:'2',style:{background:'#DCDCDC',align:'center',fontFamily:'B Nazanin'}},// merge
             {x: 17, y: 1, value: '',mergeRow:'', mergeCol:''},
             {x: 17, y: 2, value: '84',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 17, y: 3, value: '55',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 17, y: 4, value: '29',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 17, y: 5, value: '45',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 17, y: 6, value: '45 ',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 17, y: 7, value: '0',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},//
             {x: 17, y: 8, value: '7',mergeRow:'', mergeCol:'2',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 17, y: 9, value: '',mergeRow:'', mergeCol:''},
             {x: 17, y: 10, value: '223',mergeRow:'', mergeCol:'3',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 17, y: 11, value: '',mergeRow:'', mergeCol:''},
             {x: 17, y: 12, value: '',mergeRow:'', mergeCol:''},



             {x: 18, y: 0, value: 'تهران',mergeRow:'', mergeCol:'2',style:{background:'#DCDCDC',align:'center',fontFamily:'B Nazanin'}},// merge
             {x: 18, y: 1, value: '',mergeRow:'', mergeCol:''},
             {x: 18, y: 2, value: '271',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 18, y: 3, value: '212',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 18, y: 4, value: '59',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 18, y: 5, value: '396',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 18, y: 6, value: '380 ',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 18, y: 7, value: '16',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},//
             {x: 18, y: 8, value: '4',mergeRow:'', mergeCol:'2',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 18, y: 9, value: '',mergeRow:'', mergeCol:''},
             {x: 18, y: 10, value: '274',mergeRow:'', mergeCol:'3',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 18, y: 11, value: '',mergeRow:'', mergeCol:''},
             {x: 18, y: 12, value: '',mergeRow:'', mergeCol:''},




             {x: 19, y: 0, value: 'چهار محال و بختياري',mergeRow:'', mergeCol:'2',style:{background:'#DCDCDC',align:'center',fontFamily:'B Nazanin'}},// merge
             {x: 19, y: 1, value: '',mergeRow:'', mergeCol:''},
             {x: 19, y: 2, value: '111',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 19, y: 3, value: '40',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 19, y: 4, value: '71',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 19, y: 5, value: '33',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 19, y: 6, value: '26 ',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 19, y: 7, value: '7',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},//
             {x: 19, y: 8, value: '5',mergeRow:'', mergeCol:'2',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 19, y: 9, value: '',mergeRow:'', mergeCol:''},
             {x: 19, y: 10, value: '314',mergeRow:'', mergeCol:'3',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 19, y: 11, value: '',mergeRow:'', mergeCol:''},
             {x: 19, y: 12, value: '',mergeRow:'', mergeCol:''},



             {x: 20, y: 0, value: 'خراسان جنوبي',mergeRow:'', mergeCol:'2',style:{background:'#DCDCDC',align:'center',fontFamily:'B Nazanin'}},// merge
             {x: 20, y: 1, value: '',mergeRow:'', mergeCol:''},
             {x: 20, y: 2, value: '95',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 20, y: 3, value: '37',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 20, y: 4, value: '58',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 20, y: 5, value: '62',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 20, y: 6, value: '57 ',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 20, y: 7, value: '5',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},//
             {x: 20, y: 8, value: '0',mergeRow:'', mergeCol:'2',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 20, y: 9, value: '',mergeRow:'', mergeCol:''},
             {x: 20, y: 10, value: '316',mergeRow:'', mergeCol:'3',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 20, y: 11, value: '',mergeRow:'', mergeCol:''},
             {x: 20, y: 12, value: '',mergeRow:'', mergeCol:''},



             {x: 21, y: 0, value: 'خراسان رضوی',mergeRow:'', mergeCol:'2',style:{background:'#DCDCDC',align:'center',fontFamily:'B Nazanin'}},// merge
             {x: 21, y: 1, value: '',mergeRow:'', mergeCol:''},
             {x: 21, y: 2, value: '381',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 21, y: 3, value: '200',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 21, y: 4, value: '181',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 21, y: 5, value: '369',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 21, y: 6, value: '366 ',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 21, y: 7, value: '3',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},//
             {x: 21, y: 8, value: '12',mergeRow:'', mergeCol:'2',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 21, y: 9, value: '',mergeRow:'', mergeCol:''},
             {x: 21, y: 10, value: '1403',mergeRow:'', mergeCol:'3',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 21, y: 11, value: '',mergeRow:'', mergeCol:''},
             {x: 21, y: 12, value: '',mergeRow:'', mergeCol:''},



             {x: 22, y: 0, value: 'خراسان شمالي',mergeRow:'', mergeCol:'2',style:{background:'#DCDCDC',align:'center',fontFamily:'B Nazanin'}},// merge
             {x: 22, y: 1, value: '',mergeRow:'', mergeCol:''},
             {x: 22, y: 2, value: '89',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 22, y: 3, value: '28',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 22, y: 4, value: '61',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 22, y: 5, value: '58',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 22, y: 6, value: '54 ',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 22, y: 7, value: '4',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},//
             {x: 22, y: 8, value: '6',mergeRow:'', mergeCol:'2',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 22, y: 9, value: '',mergeRow:'', mergeCol:''},
             {x: 22, y: 10, value: '428',mergeRow:'', mergeCol:'3',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 22, y: 11, value: '',mergeRow:'', mergeCol:''},
             {x: 22, y: 12, value: '',mergeRow:'', mergeCol:''},



             {x: 23, y: 0, value: 'خوزستان',mergeRow:'', mergeCol:'2',style:{background:'#DCDCDC',align:'center',fontFamily:'B Nazanin'}},// merge
             {x: 23, y: 1, value: '',mergeRow:'', mergeCol:''},
             {x: 23, y: 2, value: '271',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 23, y: 3, value: '169',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 23, y: 4, value: '102',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 23, y: 5, value: '153',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 23, y: 6, value: '149 ',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 23, y: 7, value: '4',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},//
             {x: 23, y: 8, value: '14 ',mergeRow:'', mergeCol:'2',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 23, y: 9, value: '',mergeRow:'', mergeCol:''},
             {x: 23, y: 10, value: '863',mergeRow:'', mergeCol:'3',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 23, y: 11, value: '',mergeRow:'', mergeCol:''},
             {x: 23, y: 12, value: '',mergeRow:'', mergeCol:''},


             {x: 24, y: 0, value: 'زنجان',mergeRow:'', mergeCol:'2',style:{background:'#DCDCDC',align:'center',fontFamily:'B Nazanin'}},// merge
             {x: 24, y: 1, value: '',mergeRow:'', mergeCol:''},
             {x: 24, y: 2, value: '102',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 24, y: 3, value: '45',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 24, y: 4, value: '57',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 24, y: 5, value: '76',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 24, y: 6, value: '76 ',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 24, y: 7, value: '0',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},//
             {x: 24, y: 8, value: '12 ',mergeRow:'', mergeCol:'2',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 24, y: 9, value: '',mergeRow:'', mergeCol:''},
             {x: 24, y: 10, value: '448',mergeRow:'', mergeCol:'3',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 24, y: 11, value: '',mergeRow:'', mergeCol:''},
             {x: 24, y: 12, value: '',mergeRow:'', mergeCol:''},


             {x: 25, y: 0, value: 'سمنان',mergeRow:'', mergeCol:'2',style:{background:'#DCDCDC',align:'center',fontFamily:'B Nazanin'}},// merge
             {x: 25, y: 1, value: '',mergeRow:'', mergeCol:''},
             {x: 25, y: 2, value: '78 ',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 25, y: 3, value: '53',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 25, y: 4, value: '25',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 25, y: 5, value: '16',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 25, y: 6, value: '16 ',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 25, y: 7, value: '0',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},//
             {x: 25, y: 8, value: '1 ',mergeRow:'', mergeCol:'2',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 25, y: 9, value: '',mergeRow:'', mergeCol:''},
             {x: 25, y: 10, value: '136',mergeRow:'', mergeCol:'3',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 25, y: 11, value: '',mergeRow:'', mergeCol:''},
             {x: 25, y: 12, value: '',mergeRow:'', mergeCol:''},



             {x: 26, y: 0, value: 'سيستان و بلوچستان',mergeRow:'', mergeCol:'2',style:{background:'#DCDCDC',align:'center',fontFamily:'B Nazanin'}},// merge
             {x: 26, y: 1, value: '',mergeRow:'', mergeCol:''},
             {x: 26, y: 2, value: '208 ',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 26, y: 3, value: '77',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 26, y: 4, value: '131',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 26, y: 5, value: '49',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 26, y: 6, value: '46 ',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 26, y: 7, value: '3',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},//
             {x: 26, y: 8, value: '35 ',mergeRow:'', mergeCol:'2',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 26, y: 9, value: '',mergeRow:'', mergeCol:''},
             {x: 26, y: 10, value: '953',mergeRow:'', mergeCol:'3',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 26, y: 11, value: '',mergeRow:'', mergeCol:''},
             {x: 26, y: 12, value: '',mergeRow:'', mergeCol:''},



             {x: 27, y: 0, value: 'فارس',mergeRow:'', mergeCol:'2',style:{background:'#DCDCDC',align:'center',fontFamily:'B Nazanin'}},// merge
             {x: 27, y: 1, value: '',mergeRow:'', mergeCol:''},
             {x: 27, y: 2, value: '63 ',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 27, y: 3, value: '24',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 27, y: 4, value: '39',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 27, y: 5, value: '37',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 27, y: 6, value: '32 ',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 27, y: 7, value: '5',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},//
             {x: 27, y: 8, value: '7 ',mergeRow:'', mergeCol:'2',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 27, y: 9, value: '',mergeRow:'', mergeCol:''},
             {x: 27, y: 10, value: '1075',mergeRow:'', mergeCol:'3',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 27, y: 11, value: '',mergeRow:'', mergeCol:''},
             {x: 27, y: 12, value: '',mergeRow:'', mergeCol:''},



             {x: 28, y: 0, value: 'قزوین',mergeRow:'', mergeCol:'2',style:{background:'#DCDCDC',align:'center',fontFamily:'B Nazanin'}},// merge
             {x: 28, y: 1, value: '',mergeRow:'', mergeCol:''},
             {x: 28, y: 2, value: '87 ',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 28, y: 3, value: '44',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 28, y: 4, value: '43',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 28, y: 5, value: '92',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 28, y: 6, value: '92 ',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 28, y: 7, value: '0',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},//
             {x: 28, y: 8, value: '5 ',mergeRow:'', mergeCol:'2',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 28, y: 9, value: '',mergeRow:'', mergeCol:''},
             {x: 28, y: 10, value: '277',mergeRow:'', mergeCol:'3',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 28, y: 11, value: '',mergeRow:'', mergeCol:''},
             {x: 28, y: 12, value: '',mergeRow:'', mergeCol:''},



             {x: 29, y: 0, value: 'قم',mergeRow:'', mergeCol:'2',style:{background:'#DCDCDC',align:'center',fontFamily:'B Nazanin'}},// merge
             {x: 29, y: 1, value: '',mergeRow:'', mergeCol:''},
             {x: 29, y: 2, value: '56 ',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 29, y: 3, value: '47',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 29, y: 4, value: '9',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 29, y: 5, value: '144',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 29, y: 6, value: '144 ',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 29, y: 7, value: '0',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},//
             {x: 29, y: 8, value: '0 ',mergeRow:'', mergeCol:'2',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 29, y: 9, value: '',mergeRow:'', mergeCol:''},
             {x: 29, y: 10, value: '59',mergeRow:'', mergeCol:'3',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 29, y: 11, value: '',mergeRow:'', mergeCol:''},
             {x: 29, y: 12, value: '',mergeRow:'', mergeCol:''},




             {x: 30, y: 0, value: 'کردستان',mergeRow:'', mergeCol:'2',style:{background:'#DCDCDC',align:'center',fontFamily:'B Nazanin'}},// merge
             {x: 30, y: 1, value: '',mergeRow:'', mergeCol:''},
             {x: 30, y: 2, value: '148 ',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 30, y: 3, value: '75 ',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 30, y: 4, value: '73',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 30, y: 5, value: '107',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 30, y: 6, value: '106 ',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 30, y: 7, value: '1',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},//
             {x: 30, y: 8, value: '3 ',mergeRow:'', mergeCol:'2',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 30, y: 9, value: '',mergeRow:'', mergeCol:''},
             {x: 30, y: 10, value: '618',mergeRow:'', mergeCol:'3',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 30, y: 11, value: '',mergeRow:'', mergeCol:''},
             {x: 30, y: 12, value: '',mergeRow:'', mergeCol:''},




             {x: 31, y: 0, value: 'کرمان',mergeRow:'', mergeCol:'2',style:{background:'#DCDCDC',align:'center',fontFamily:'B Nazanin'}},// merge
             {x: 31, y: 1, value: '',mergeRow:'', mergeCol:''},
             {x: 31, y: 2, value: '203 ',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 31, y: 3, value: '90 ',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 31, y: 4, value: '113',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 31, y: 5, value: '114',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 31, y: 6, value: '108 ',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 31, y: 7, value: '6',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},//
             {x: 31, y: 8, value: '14 ',mergeRow:'', mergeCol:'2',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 31, y: 9, value: '',mergeRow:'', mergeCol:''},
             {x: 31, y: 10, value: '855',mergeRow:'', mergeCol:'3',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 31, y: 11, value: '',mergeRow:'', mergeCol:''},
             {x: 31, y: 12, value: '',mergeRow:'', mergeCol:''},



             {x: 32, y: 0, value: 'کرمانشاه',mergeRow:'', mergeCol:'2',style:{background:'#DCDCDC',align:'center',fontFamily:'B Nazanin'}},// merge
             {x: 32, y: 1, value: '',mergeRow:'', mergeCol:''},
             {x: 32, y: 2, value: '130 ',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 32, y: 3, value: '69 ',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 32, y: 4, value: '61',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 32, y: 5, value: '25',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 32, y: 6, value: '25 ',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 32, y: 7, value: '0',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},//
             {x: 32, y: 8, value: '12 ',mergeRow:'', mergeCol:'2',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 32, y: 9, value: '',mergeRow:'', mergeCol:''},
             {x: 32, y: 10, value: '649',mergeRow:'', mergeCol:'3',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 32, y: 11, value: '',mergeRow:'', mergeCol:''},
             {x: 32, y: 12, value: '',mergeRow:'', mergeCol:''},




             {x: 33, y: 0, value: 'كهگيلويه و بويراحمد',mergeRow:'', mergeCol:'2',style:{background:'#DCDCDC',align:'center',fontFamily:'B Nazanin'}},// merge
             {x: 33, y: 1, value: '',mergeRow:'', mergeCol:''},
             {x: 33, y: 2, value: '69 ',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 33, y: 3, value: '27 ',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 33, y: 4, value: '42',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 33, y: 5, value: '51',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 33, y: 6, value: '49 ',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 33, y: 7, value: '2',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},//
             {x: 33, y: 8, value: '11 ',mergeRow:'', mergeCol:'2',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 33, y: 9, value: '',mergeRow:'', mergeCol:''},
             {x: 33, y: 10, value: '345',mergeRow:'', mergeCol:'3',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 33, y: 11, value: '',mergeRow:'', mergeCol:''},
             {x: 33, y: 12, value: '',mergeRow:'', mergeCol:''},




             {x: 34, y: 0, value: 'گلستان',mergeRow:'', mergeCol:'2',style:{background:'#DCDCDC',align:'center',fontFamily:'B Nazanin'}},// merge
             {x: 34, y: 1, value: '',mergeRow:'', mergeCol:''},
             {x: 34, y: 2, value: '144 ',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 34, y: 3, value: '47 ',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 34, y: 4, value: '97',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 34, y: 5, value: '38',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 34, y: 6, value: '4 ',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 34, y: 7, value: '34',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},//
             {x: 34, y: 8, value: '6 ',mergeRow:'', mergeCol:'2',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 34, y: 9, value: '',mergeRow:'', mergeCol:''},
             {x: 34, y: 10, value: '608',mergeRow:'', mergeCol:'3',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 34, y: 11, value: '',mergeRow:'', mergeCol:''},
             {x: 34, y: 12, value: '',mergeRow:'', mergeCol:''},


             {x: 35, y: 0, value: 'گیلان',mergeRow:'', mergeCol:'2',style:{background:'#DCDCDC',align:'center',fontFamily:'B Nazanin'}},// merge
             {x: 35, y: 1, value: '',mergeRow:'', mergeCol:''},
             {x: 35, y: 2, value: '192 ',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 35, y: 3, value: '96 ',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 35, y: 4, value: '96',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 35, y: 5, value: '44',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 35, y: 6, value: '43 ',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 35, y: 7, value: '1',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},//
             {x: 35, y: 8, value: '3',mergeRow:'', mergeCol:'2',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 35, y: 9, value: '',mergeRow:'', mergeCol:''},
             {x: 35, y: 10, value: '971',mergeRow:'', mergeCol:'3',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 35, y: 11, value: '',mergeRow:'', mergeCol:''},
             {x: 35, y: 12, value: '',mergeRow:'', mergeCol:''},




             {x: 36, y: 0, value: 'لرستان',mergeRow:'', mergeCol:'2',style:{background:'#DCDCDC',align:'center',fontFamily:'B Nazanin'}},// merge
             {x: 36, y: 1, value: '',mergeRow:'', mergeCol:''},
             {x: 36, y: 2, value: '152',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 36, y: 3, value: '73  ',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 36, y: 4, value: '5',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 36, y: 5, value: '44 ',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 36, y: 6, value: '41 ',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 36, y: 7, value: '3',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},//
             {x: 36, y: 8, value: '2',mergeRow:'', mergeCol:'2',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 36, y: 9, value: '',mergeRow:'', mergeCol:''},
             {x: 36, y: 10, value: '642',mergeRow:'', mergeCol:'3',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 36, y: 11, value: '',mergeRow:'', mergeCol:''},
             {x: 36, y: 12, value: '',mergeRow:'', mergeCol:''},




             {x: 37, y: 0, value: 'مازندران ',mergeRow:'', mergeCol:'2',style:{background:'#DCDCDC',align:'center',fontFamily:'B Nazanin'}},// merge
             {x: 37, y: 1, value: '',mergeRow:'', mergeCol:''},
             {x: 37, y: 2, value: '292',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 37, y: 3, value: '96  ',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 37, y: 4, value: '1 ',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 37, y: 5, value: '76 ',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 37, y: 6, value: '74 ',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 37, y: 7, value: '2',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},//
             {x: 37, y: 8, value: '10',mergeRow:'', mergeCol:'2',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 37, y: 9, value: '',mergeRow:'', mergeCol:''},
             {x: 37, y: 10, value: '1288',mergeRow:'', mergeCol:'3',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 37, y: 11, value: '',mergeRow:'', mergeCol:''},
             {x: 37, y: 12, value: '',mergeRow:'', mergeCol:''},




             {x: 38, y: 0, value: 'مرکزی ',mergeRow:'', mergeCol:'2',style:{background:'#DCDCDC',align:'center',fontFamily:'B Nazanin'}},// merge
             {x: 38, y: 1, value: '',mergeRow:'', mergeCol:''},
             {x: 38, y: 2, value: '22',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 38, y: 3, value: '14   ',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 38, y: 4, value: '1 ',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 38, y: 5, value: '22 ',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 38, y: 6, value: '22 ',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 38, y: 7, value: '0',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},//
             {x: 38, y: 8, value: '0',mergeRow:'', mergeCol:'2',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 38, y: 9, value: '',mergeRow:'', mergeCol:''},
             {x: 38, y: 10, value: '396',mergeRow:'', mergeCol:'3',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 38, y: 11, value: '',mergeRow:'', mergeCol:''},
             {x: 38, y: 12, value: '',mergeRow:'', mergeCol:''},


             {x: 39, y: 0, value: 'هرمزگان ',mergeRow:'', mergeCol:'2',style:{background:'#DCDCDC',align:'center',fontFamily:'B Nazanin'}},// merge
             {x: 39, y: 1, value: '',mergeRow:'', mergeCol:''},
             {x: 39, y: 2, value: '154',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 39, y: 3, value: '59   ',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 39, y: 4, value: '1 ',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 39, y: 5, value: '50 ',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 39, y: 6, value: '45 ',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 39, y: 7, value: '5',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},//
             {x: 39, y: 8, value: '30',mergeRow:'', mergeCol:'2',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 39, y: 9, value: '',mergeRow:'', mergeCol:''},
             {x: 39, y: 10, value: '556',mergeRow:'', mergeCol:'3',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 39, y: 11, value: '',mergeRow:'', mergeCol:''},
             {x: 39, y: 12, value: '',mergeRow:'', mergeCol:''},


             {x: 40, y: 0, value: 'هرمزگان ',mergeRow:'', mergeCol:'2',style:{background:'#DCDCDC',align:'center',fontFamily:'B Nazanin'}},// merge
             {x: 40, y: 1, value: '',mergeRow:'', mergeCol:''},
             {x: 40, y: 2, value: '152',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 40, y: 3, value: '66   ',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 40, y: 4, value: '1 ',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 40, y: 5, value: '43 ',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 40, y: 6, value: '32 ',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 40, y: 7, value: '11',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},//
             {x: 40, y: 8, value: '5',mergeRow:'', mergeCol:'2',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 40, y: 9, value: '',mergeRow:'', mergeCol:''},
             {x: 40, y: 10, value: '575',mergeRow:'', mergeCol:'3',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 40, y: 11, value: '',mergeRow:'', mergeCol:''},
             {x: 40, y: 12, value: '',mergeRow:'', mergeCol:''},




             {x: 41, y: 0, value: 'یزد ',mergeRow:'', mergeCol:'2',style:{background:'#DCDCDC',align:'center',fontFamily:'B Nazanin'}},// merge
             {x: 41, y: 1, value: '',mergeRow:'', mergeCol:''},
             {x: 41, y: 2, value: '85',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 41, y: 3, value: '49   ',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 41, y: 4, value: '1 ',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 41, y: 5, value: '91 ',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 41, y: 6, value: '91 ',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 41, y: 7, value: '0',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},//
             {x: 41, y: 8, value: '5',mergeRow:'', mergeCol:'2',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 41, y: 9, value: '',mergeRow:'', mergeCol:''},
             {x: 41, y: 10, value: '171',mergeRow:'', mergeCol:'3',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 41, y: 11, value: '',mergeRow:'', mergeCol:''},
             {x: 41, y: 12, value: '',mergeRow:'', mergeCol:''},



             {x: 42, y: 0, value: '1395 ',mergeRow:'', mergeCol:'2',style:{background:'#DCDCDC',align:'center',fontFamily:'B Nazanin'}},// merge
             {x: 42, y: 1, value: '',mergeRow:'', mergeCol:''},
             {x: 42, y: 2, value: '//',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 42, y: 3, value: '×   ',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 42, y: 4, value: '×× ',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 42, y: 5, value: ' ',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 42, y: 6, value: ' ',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 42, y: 7, value: '',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},//
             {x: 42, y: 8, value: '',mergeRow:'', mergeCol:'2',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 42, y: 9, value: '',mergeRow:'', mergeCol:''},
             {x: 42, y: 10, value: '',mergeRow:'', mergeCol:'3',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 42, y: 11, value: '',mergeRow:'', mergeCol:''},
             {x: 42, y: 12, value: '',mergeRow:'', mergeCol:''}


         ];// data / داده مربوط به مراكز ارائه دهنده مراقبت هاي اوليه بهداشتي //


         let docx = new Docx('temp2.docx', 'outputUnitTestTmp/');
         docx.createP();
         docx.addContentP('٦-١٨- مراكز ارائه دهنده مراقبت هاي اوليه بهداشتي'  ,{fontSize:12});
         docx.createTable(data);
         docx.createP();
         docx.createP();
         docx.addContentP('مأخذ - وزارت بهداشت، درمان و آموزش پزشكي.'  ,{fontSize:8});
         docx.createP();
         docx.addContentP('- وزارت بهداشت، درمان و آموزش پزشكي. دفتر مديريت آمار و فناوري اطلاعات.'  ,{fontSize:8});

         let out = docx.generate();
         if(out == false){
             fail("Don't create File of docx");
         }else{
             console.log("Create File of Docx");
             done();
         }


     });*/ // it  /// فایل مربوط به مراكز ارائه دهنده مراقبت هاي اوليه بهداشتي
    /*  it('Template Three', function(done){///  فایل مربوط به انواع مرسولات پستی صادر شده برون استانی
 
         let data = [
             {x: 1, y: 0, value: '',mergeRow:'2', mergeCol:'3',style:{background:'#f7f7f7',align:'center',fontFamily:'B Nazanin'}},   //mergeRow:(x) ,,,, mergeCol:(y)
             {x: 1, y: 1, value: '', mergeRow:'', mergeCol:''},
             {x: 1, y: 2, value: '',mergeRow:'', mergeCol:''},
             {x: 1, y: 3, value: 'جمع',mergeRow:'2', mergeCol:'',style:{background:'#f7f7f7',align:'center',fontFamily:'B Nazanin'}},
             {x: 1, y: 4, value: 'عادی',mergeRow:'2', mergeCol:'',style:{background:'#f7f7f7',align:'center',fontFamily:'B Nazanin'}},
             {x: 1, y: 5, value: 'سفارشی',mergeRow:'', mergeCol:'4',style:{background:'#f7f7f7',align:'center',fontFamily:'B Nazanin'}},
             {x: 1, y: 6, value: '',mergeRow:'', mergeCol:''},
             {x: 1, y: 7, value: '',mergeRow:'', mergeCol:''},
             {x: 1, y: 8, value: '',mergeRow:'', mergeCol:''},
             {x: 1, y: 9, value: 'پیشتاز',mergeRow:'', mergeCol:'2',style:{background:'#f7f7f7',align:'center',fontFamily:'B Nazanin'}},
             {x: 1, y: 10, value: '',mergeRow:'', mergeCol:''},
             {x: 1, y: 11, value: 'پست ویژه',mergeRow:'', mergeCol:'2',style:{background:'#f7f7f7',align:'center',fontFamily:'B Nazanin'}},
             {x: 1, y: 12, value: '',mergeRow:'', mergeCol:''},
 
 
 
 
             {x: 2, y: 0, value: '',mergeRow:'', mergeCol:''},
             {x: 2, y: 1, value: '',mergeRow:'', mergeCol:''},
             {x: 2, y: 2, value: '',mergeRow:'', mergeCol:''},
             {x: 2, y: 3, value: '',mergeRow:'', mergeCol:''},
             {x: 2, y: 4, value: '',mergeRow:'', mergeCol:''},
             {x: 2, y: 5, value: 'نامه',mergeRow:'', mergeCol:'',style:{background:'#f7f7f7',align:'center',fontFamily:'B Nazanin'}},
             {x: 2, y: 6, value: 'بسته وکیسه',mergeRow:'', mergeCol:'',style:{background:'#f7f7f7',align:'center',fontFamily:'B Nazanin'}},
             {x: 2, y: 7, value: 'مطبوعات',mergeRow:'', mergeCol:'',style:{background:'#f7f7f7',align:'center',fontFamily:'B Nazanin'}},
             {x: 2, y: 8, value: 'امانت',mergeRow:'', mergeCol:'',style:{background:'#f7f7f7',align:'center',fontFamily:'B Nazanin'}},
             {x: 2, y: 9, value: 'پاکت',mergeRow:'', mergeCol:'',style:{background:'#f7f7f7',align:'center',fontFamily:'B Nazanin'}},
             {x: 2, y: 10, value: 'بسته',mergeRow:'', mergeCol:'',style:{background:'#f7f7f7',align:'center',fontFamily:'B Nazanin'}},
             {x: 2, y: 11, value: 'پاکت',mergeRow:'', mergeCol:'',style:{background:'#f7f7f7',align:'center',fontFamily:'B Nazanin'}},
             {x: 2, y: 12, value: 'بسته',mergeRow:'', mergeCol:'',style:{background:'#f7f7f7',align:'center',fontFamily:'B Nazanin'}},
 
 
 
 
 
             {x: 3, y: 0, value: 'سال',mergeRow:'24', mergeCol:'',style:{background:'#f7f7f7',align:'center',fontFamily:'B Nazanin'}},
             {x: 3, y: 1, value: '1375',mergeRow:'', mergeCol:'2',style:{background:'#f7f7f7',align:'center',fontFamily:'B Nazanin'}},
             {x: 3, y: 2, value: '',mergeRow:'', mergeCol:''},
             {x: 3, y: 3, value: '0',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 3, y: 4, value: '0',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 3, y: 5, value: '0',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 3, y: 6, value: '0',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 3, y: 7, value: '0',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 3, y: 8, value: '0',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 3, y: 9, value: '0',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 3, y: 10, value: '0',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 3, y: 11, value: '0',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 3, y: 12, value: '0',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             //
             //
             //
             {x: 4, y: 0, value: '',mergeRow:'', mergeCol:''},
             {x: 4, y: 1, value: '1380',mergeRow:'', mergeCol:'2',style:{background:'#f7f7f7',align:'center',fontFamily:'B Nazanin'}},
             {x: 4, y: 2, value: '',mergeRow:'', mergeCol:''},
             {x: 4, y: 3, value: '0',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 4, y: 4, value: '0',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 4, y: 5, value: '0',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 4, y: 6, value: '0',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 4, y: 7, value: '0',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 4, y: 8, value: '0',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 4, y: 9, value: '0',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 4, y: 10, value: '0',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 4, y: 11, value: '0',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 4, y: 12, value: '0',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             //
             //
             //
             // //
             {x: 5, y: 0, value: '',mergeRow:'', mergeCol:''},
             {x: 5, y: 1, value: '1385',mergeRow:'', mergeCol:'2',style:{background:'#f7f7f7',align:'center',fontFamily:'B Nazanin'}},
             {x: 5, y: 2, value: '',mergeRow:'', mergeCol:''},
             {x: 5, y: 3, value: '0',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 5, y: 4, value: '0',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 5, y: 5, value: '0',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 5, y: 6, value: '0',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 5, y: 7, value: '0',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 5, y: 8, value: '0',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 5, y: 9, value: '0',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 5, y: 10, value: '0',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 5, y: 11, value: '0',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 5, y: 12, value: '0',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             //
             // //
             // //
             // //
             // //
             {x: 6, y: 0, value: '',mergeRow:'', mergeCol:''},
             {x: 6, y: 1, value: '1389',mergeRow:'', mergeCol:'2',style:{background:'#f7f7f7',align:'center',fontFamily:'B Nazanin'}},
             {x: 6, y: 2, value: '',mergeRow:'', mergeCol:''},
             {x: 6, y: 3, value: '0',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 6, y: 4, value: '0',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 6, y: 5, value: '0',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 6, y: 6, value: '0',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 6, y: 7, value: '0',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 6, y: 8, value: '0',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 6, y: 9, value: '0',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 6, y: 10, value: '0',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 6, y: 11, value: '0',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 6, y: 12, value: '0',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             //
             // //
             // //
             // //
             {x: 7, y: 0, value: '',mergeRow:'', mergeCol:''},
             {x: 7, y: 1, value: '1390',mergeRow:'', mergeCol:'2',style:{background:'#f7f7f7',align:'center',fontFamily:'B Nazanin'}},
             {x: 7, y: 2, value: '',mergeRow:'', mergeCol:''},
             {x: 7, y: 3, value: '8457868',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 7, y: 4, value: '8381192',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 7, y: 5, value: '12746',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 7, y: 6, value: '388',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 7, y: 7, value: '129',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 7, y: 8, value: '3533',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 7, y: 9, value: '55548',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 7, y: 10, value: '3219',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 7, y: 11, value: '957',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 7, y: 12, value: '156',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             // //
             // //
             // //
             // //
             // //
             // //
             {x: 8, y: 0, value: '',mergeRow:'', mergeCol:''},
             {x: 8, y: 1, value: '1391',mergeRow:'', mergeCol:'2',style:{background:'#f7f7f7',align:'center',fontFamily:'B Nazanin'}},
             {x: 8, y: 2, value: '',mergeRow:'', mergeCol:''},
             {x: 8, y: 3, value: '7593734',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 8, y: 4, value: '6821280',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 8, y: 5, value: '108135',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 8, y: 6, value: '6088',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 8, y: 7, value: '1124',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 8, y: 8, value: '37833',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 8, y: 9, value: '561737',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 8, y: 10, value: '41107',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 8, y: 11, value: '14680',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 8, y: 12, value: '1750',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             // //
             // //
             // //
             {x: 9, y: 0, value: '',mergeRow:'', mergeCol:''},
             {x: 9, y: 1, value: '1392',mergeRow:'', mergeCol:'2',style:{background:'#f7f7f7',align:'center',fontFamily:'B Nazanin'}},
             {x: 9, y: 2, value: '',mergeRow:'', mergeCol:''},
             {x: 9, y: 3, value: '5813650',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 9, y: 4, value: '5100922',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 9, y: 5, value: '109268',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 9, y: 6, value: '7040',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 9, y: 7, value: '819',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 9, y: 8, value: '30849',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 9, y: 9, value: '497955',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 9, y: 10, value: '52560',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 9, y: 11, value: '12697',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 9, y: 12, value: '1540',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             // //
             // //
             // //
             // //
             //
             {x: 10, y: 0, value: '',mergeRow:'', mergeCol:''},
             {x: 10, y: 1, value: '1393',mergeRow:'', mergeCol:'2',style:{background:'#f7f7f7',align:'center',fontFamily:'B Nazanin'}},
             {x: 10, y: 2, value: '',mergeRow:'', mergeCol:''},
             {x: 10, y: 3, value: '7259945',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 10, y: 4, value: '6591227',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 10, y: 5, value: '82854',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 10, y: 6, value: '8444',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 10, y: 7, value: '939',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 10, y: 8, value: '27623',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 10, y: 9, value: '480544',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 10, y: 10, value: '49151',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 10, y: 11, value: '15794',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 10, y: 12, value: '3369',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
 
 
 
             {x: 11, y: 0, value: 'شهرستان',mergeRow:'6', mergeCol:'',style:{background:'#f7f7f7',align:'center',fontFamily:'B Nazanin'}},
             {x: 11, y: 1, value: 'آستارا',mergeRow:'', mergeCol:'2',style:{background:'#f7f7f7',align:'center',fontFamily:'B Nazanin'}},
             {x: 11, y: 2, value: '',mergeRow:'', mergeCol:''},
             {x: 11, y: 3, value: '183614',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 11, y: 4, value: '147809',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 11, y: 5, value: '7031',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 11, y: 6, value: '327',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 11, y: 7, value: '22',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 11, y: 8, value: '6331',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 11, y: 9, value: '19160',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 11, y: 10, value: '2810',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 11, y: 11, value: '119',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 11, y: 12, value: '5',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             //
             // //
             // //
             // //
             {x: 12, y: 0, value: '',mergeRow:'', mergeCol:''},
             {x: 12, y: 1, value: 'آستانه اشرفیه',mergeRow:'', mergeCol:'2',style:{background:'#f7f7f7',align:'center',fontFamily:'B Nazanin'}},
             {x: 12, y: 2, value: '',mergeRow:'', mergeCol:''},
             {x: 12, y: 3, value: '144109',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 12, y: 4, value: '125266',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 12, y: 5, value: '2412',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 12, y: 6, value: '235',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 12, y: 7, value: '3',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 12, y: 8, value: '499',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 12, y: 9, value: '13878',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 12, y: 10, value: '1447',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 12, y: 11, value: '334',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 12, y: 12, value: '35',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             // //
             // //
             // //
             {x: 13, y: 0, value: '',mergeRow:'', mergeCol:''},
             {x: 13, y: 1, value: 'املش',mergeRow:'', mergeCol:'2',style:{background:'#f7f7f7',align:'center',fontFamily:'B Nazanin'}},
             {x: 13, y: 2, value: '',mergeRow:'', mergeCol:''},
             {x: 13, y: 3, value: '76878',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 13, y: 4, value: '71419',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 13, y: 5, value: '649',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 13, y: 6, value: '192',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 13, y: 7, value: '0',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 13, y: 8, value: '439',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 13, y: 9, value: '3447',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 13, y: 10, value: '646',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 13, y: 11, value: '81',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 13, y: 12, value: '5',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
 
             // //
             //
             //
             //
             //
             {x: 14, y: 0, value: '',mergeRow:'', mergeCol:''},
             {x: 14, y: 1, value: 'بندر انزلی',mergeRow:'', mergeCol:'2',style:{background:'#f7f7f7',align:'center',fontFamily:'B Nazanin'}},
             {x: 14, y: 2, value: '',mergeRow:'', mergeCol:''},
             {x: 14, y: 3, value: '607510',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 14, y: 4, value: '559994',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 14, y: 5, value: '3386',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 14, y: 6, value: '860',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 14, y: 7, value: '2',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 14, y: 8, value: '1295',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 14, y: 9, value: '36741',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 14, y: 10, value: '3986',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 14, y: 11, value: '1178',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 14, y: 12, value: '68',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             //
             //
             //
             {x: 15, y: 0, value: '',mergeRow:'', mergeCol:''},
             {x: 15, y: 1, value: 'تالش',mergeRow:'', mergeCol:'2',style:{background:'#f7f7f7',align:'center',fontFamily:'B Nazanin'}},
             {x: 15, y: 2, value: '',mergeRow:'', mergeCol:''},
             {x: 15, y: 3, value: '194696',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 15, y: 4, value: '169920',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 15, y: 5, value: '2771',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 15, y: 6, value: '282',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 15, y: 7, value: '422',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 15, y: 8, value: '600',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 15, y: 9, value: '18413',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 15, y: 10, value: '1921',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 15, y: 11, value: '324',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 15, y: 12, value: '43',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             //
             //
             //
             //
             {x: 16, y: 0, value: '',mergeRow:'', mergeCol:''},
             {x: 16, y: 1, value: 'رشت',mergeRow:'', mergeCol:'2',style:{background:'#f7f7f7',align:'center',fontFamily:'B Nazanin'}},
             {x: 16, y: 2, value: '',mergeRow:'', mergeCol:''},
             {x: 16, y: 3, value: '3797315',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 16, y: 4, value: '3438022',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 16, y: 5, value: '47011',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 16, y: 6, value: '3743',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 16, y: 7, value: '479',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 16, y: 8, value: '8434',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 16, y: 9, value: '264729',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 16, y: 10, value: '23295',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 16, y: 11, value: '8896',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 16, y: 12, value: '2706',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             //
             //
             //
             //
             {x: 17, y: 0, value: '',mergeRow:'', mergeCol:''},
             {x: 17, y: 1, value: 'رضوانشهر',mergeRow:'', mergeCol:'2',style:{background:'#f7f7f7',align:'center',fontFamily:'B Nazanin'}},
             {x: 17, y: 2, value: '',mergeRow:'', mergeCol:''},
             {x: 17, y: 3, value: '117525',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 17, y: 4, value: '106026',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 17, y: 5, value: '662',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 17, y: 6, value: '52',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 17, y: 7, value: '0',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 17, y: 8, value: '288',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 17, y: 9, value: '9238',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 17, y: 10, value: '828',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 17, y: 11, value: '413',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 17, y: 12, value: '18',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             //
             //
             //
             {x: 18, y: 0, value: '',mergeRow:'', mergeCol:''},
             {x: 18, y: 1, value: 'رودبار',mergeRow:'', mergeCol:'2',style:{background:'#f7f7f7',align:'center',fontFamily:'B Nazanin'}},
             {x: 18, y: 2, value: '',mergeRow:'', mergeCol:''},
             {x: 18, y: 3, value: '511862',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 18, y: 4, value: '495348',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 18, y: 5, value: '2532',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 18, y: 6, value: '257',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 18, y: 7, value: '1',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 18, y: 8, value: '405',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 18, y: 9, value: '11136',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 18, y: 10, value: '1746',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 18, y: 11, value: '410',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 18, y: 12, value: '27',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             //
             //
             //
             {x: 19, y: 0, value: '',mergeRow:'', mergeCol:''},
             {x: 19, y: 1, value: 'رودسر',mergeRow:'', mergeCol:'2',style:{background:'#f7f7f7',align:'center',fontFamily:'B Nazanin'}},
             {x: 19, y: 2, value: '',mergeRow:'', mergeCol:''},
             {x: 19, y: 3, value: '227333',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 19, y: 4, value: '198359',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 19, y: 5, value: '2748',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 19, y: 6, value: '550 ',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 19, y: 7, value: '5',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 19, y: 8, value: '2013',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 19, y: 9, value: '19562',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 19, y: 10, value: '2750',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 19, y: 11, value: '1263',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 19, y: 12, value: '83',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             //
             //
             //
             //
             {x: 20, y: 0, value: '',mergeRow:'', mergeCol:''},
             {x: 20, y: 1, value: 'سیاهکل',mergeRow:'', mergeCol:'2',style:{background:'#f7f7f7',align:'center',fontFamily:'B Nazanin'}},
             {x: 20, y: 2, value: '',mergeRow:'', mergeCol:''},
             {x: 20, y: 3, value: '44905',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 20, y: 4, value: '40053',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 20, y: 5, value: '223',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 20, y: 6, value: '5 ',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 20, y: 7, value: '0',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 20, y: 8, value: '156',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 20, y: 9, value: '3850',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 20, y: 10, value: '419',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 20, y: 11, value: '190',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 20, y: 12, value: '9',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             //
             //
             //
             //
             {x: 21, y: 0, value: '',mergeRow:'', mergeCol:''},
             {x: 21, y: 1, value: 'شفت',mergeRow:'', mergeCol:'2',style:{background:'#f7f7f7',align:'center',fontFamily:'B Nazanin'}},
             {x: 21, y: 2, value: '',mergeRow:'', mergeCol:''},
             {x: 21, y: 3, value: '30795',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 21, y: 4, value: '27548',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 21, y: 5, value: '231',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 21, y: 6, value: '19 ',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 21, y: 7, value: '0',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 21, y: 8, value: '67',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 21, y: 9, value: '2690',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 21, y: 10, value: '204',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 21, y: 11, value: '35',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 21, y: 12, value: '1',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             //
             //
             //
             {x: 22, y: 0, value: '',mergeRow:'', mergeCol:''},
             {x: 22, y: 1, value: 'صومعه سرا',mergeRow:'', mergeCol:'2',style:{background:'#f7f7f7',align:'center',fontFamily:'B Nazanin'}},
             {x: 22, y: 2, value: '',mergeRow:'', mergeCol:''},
             {x: 22, y: 3, value: '141442',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 22, y: 4, value: '124991',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 22, y: 5, value: '748',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 22, y: 6, value: '0',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 22, y: 7, value: '191',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 22, y: 8, value: '413',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 22, y: 9, value: '13286',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 22, y: 10, value: '1524',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 22, y: 11, value: '276',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 22, y: 12, value: '13',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
 
             //
             //
             //
             {x: 23, y: 0, value: '',mergeRow:'', mergeCol:''},
             {x: 23, y: 1, value: 'فومن',mergeRow:'', mergeCol:'2',style:{background:'#f7f7f7',align:'center',fontFamily:'B Nazanin'}},
             {x: 23, y: 2, value: '',mergeRow:'', mergeCol:''},
             {x: 23, y: 3, value: '82171',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 23, y: 4, value: '68969',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 23, y: 5, value: '1228 ',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 23, y: 6, value: '195',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 23, y: 7, value: '2',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 23, y: 8, value: '664',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 23, y: 9, value: '9343',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 23, y: 10, value: '944 ',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 23, y: 11, value: '766',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 23, y: 12, value: '60',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             //
             //
             //
             {x: 24, y: 0, value: '',mergeRow:'', mergeCol:''},
             {x: 24, y: 1, value: 'لاهیجان',mergeRow:'', mergeCol:'2',style:{background:'#f7f7f7',align:'center',fontFamily:'B Nazanin'}},
             {x: 24, y: 2, value: '',mergeRow:'', mergeCol:''},
             {x: 24, y: 3, value: '830229',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 24, y: 4, value: '781260',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 24, y: 5, value: '9024',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 24, y: 6, value: '1062',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 24, y: 7, value: '2',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 24, y: 8, value: '4744',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 24, y: 9, value: '29979',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 24, y: 10, value: '3453',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 24, y: 11, value: '622',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 24, y: 12, value: '83',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
 
 
 
             {x: 25, y: 0, value: '',mergeRow:'', mergeCol:''},
             {x: 25, y: 1, value: 'لنگرود',mergeRow:'', mergeCol:'2',style:{background:'#f7f7f7',align:'center',fontFamily:'B Nazanin'}},
             {x: 25, y: 2, value: '',mergeRow:'', mergeCol:''},
             {x: 25, y: 3, value: '212913',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 25, y: 4, value: '185303',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 25, y: 5, value: '1681',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 25, y: 6, value: '467',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 25, y: 7, value: '1',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 25, y: 8, value: '1113',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 25, y: 9, value: '21075',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 25, y: 10, value: '2484',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 25, y: 11, value: '626',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 25, y: 12, value: '163',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
 
 
 
             {x: 26, y: 0, value: '',mergeRow:'', mergeCol:''},
             {x: 26, y: 1, value: 'ماسال',mergeRow:'', mergeCol:'2',style:{background:'#f7f7f7',align:'center',fontFamily:'B Nazanin'}},
             {x: 26, y: 2, value: '',mergeRow:'', mergeCol:''},
             {x: 26, y: 3, value: '56648',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 26, y: 4, value: '50940',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 26, y: 5, value: '517',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 26, y: 6, value: '7',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 26, y: 7, value: '0',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 26, y: 8, value: '162',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 26, y: 9, value: '4017',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 26, y: 10, value: '694',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 26, y: 11, value: '261',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 26, y: 12, value: '50',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
 
 
 
         ];// data / داده مربوط به انواع مرسولات پستی صادر شده برون استانی //
 
 
 
         let docx = new Docx('temp3.docx', 'outputUnitTestTmp/');
         docx.createP();
         docx.addContentP('انواع مرسولات پستی صادر شده برون استانی'  ,{fontSize:12});
         docx.createTable(data);
         docx.createP();
         docx.createP();
         docx.addContentP('مأخذ – اداره كل پست استان گيلان.'  ,{fontSize:10});
         let out = docx.generate();
         if(out == false){
             fail("Don't create File of docx");
         }else{
             console.log("Create File of Docx");
             done();
         }
 
 
     });*/ // it فایل مربوط به انواع مرسولات پستی صادر شده برون استانی/
    /* it('template Three',  function(done){ //    فایل مربوط به تفکیک جمعیت شهرستان   ///
 
         let data = [
             {x: 1, y: 0, value: '',mergeRow:'', mergeCol:'',style:{background:'#f7f7f7',align:'center',fontFamily:'B Nazanin',borderSize:15}},   //mergeRow:(x) ,,,, mergeCol:(y)
             {x: 1, y: 1, value: 'سال 1390', mergeRow:'', mergeCol:'',style:{background:'#f7f7f7',align:'center',fontFamily:'B Nazanin',borderSize:15}},
             {x: 1, y: 2, value: 'سال 1391',mergeRow:'', mergeCol:'',style:{background:'#f7f7f7',align:'center',fontFamily:'B Nazanin',borderSize:15}},
             {x: 1, y: 3, value: 'سال 1392',mergeRow:'', mergeCol:'',style:{background:'#f7f7f7',align:'center',fontFamily:'B Nazanin',borderSize:15}},
             {x: 1, y: 4, value: 'سال 1393',mergeRow:'', mergeCol:'',style:{background:'#f7f7f7',align:'center',fontFamily:'B Nazanin',borderSize:15}},
             {x: 1, y: 5, value: 'سال 1394',mergeRow:'', mergeCol:'',style:{background:'#f7f7f7',align:'center',fontFamily:'B Nazanin',borderSize:15}},
             {x: 1, y: 6, value: 'سال1395',mergeRow:'', mergeCol:'',style:{background:'#f7f7f7',align:'center',fontFamily:'B Nazanin',borderSize:15}},
 
 
             {x: 2, y: 0, value: 'کل',mergeRow:'', mergeCol:'',style:{background:'#f7f7f7',align:'center',fontFamily:'B Nazanin',borderSize:15}},   //mergeRow:(x) ,,,, mergeCol:(y)
             {x: 2, y: 1, value: '21545288', mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 2, y: 2, value: '85487525',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 2, y: 3, value: '2215659',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 2, y: 4, value: '2211255',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 2, y: 5, value: '2255555',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 2, y: 6, value: '985412',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
 
 
             {x: 3, y: 0, value: 'البرز',mergeRow:'', mergeCol:'',style:{background:'#f7f7f7',align:'center',fontFamily:'B Nazanin',borderSize:15}},   //mergeRow:(x) ,,,, mergeCol:(y)
             {x: 3, y: 1, value: '2521', mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 3, y: 2, value: '5485',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 3, y: 3, value: '514 ',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 3, y: 4, value: '6452',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 3, y: 5, value: '56588',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 3, y: 6, value: '2541',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
 
 
             {x: 4, y: 0, value: 'بندرعباس',mergeRow:'', mergeCol:'',style:{background:'#f7f7f7',align:'center',fontFamily:'B Nazanin',borderSize:15}},   //mergeRow:(x) ,,,, mergeCol:(y)
             {x: 4, y: 1, value: '145214', mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 4, y: 2, value: '2255',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 4, y: 3, value: '225552',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 4, y: 4, value: '5552',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 4, y: 5, value: '2558 ',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 4, y: 6, value: '10255',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
 
 
 
             {x: 5, y: 0, value: 'گیلان',mergeRow:'', mergeCol:'',style:{background:'#f7f7f7',align:'center',fontFamily:'B Nazanin',borderSize:15}},   //mergeRow:(x) ,,,, mergeCol:(y)
             {x: 5, y: 1, value: '25487', mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 5, y: 2, value: '32145',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 5, y: 3, value: '25555',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 5, y: 4, value: '22123',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 5, y: 5, value: '6542 ',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 5, y: 6, value: '5488',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
 
 
 
             {x: 6, y: 0, value: 'گلستان',mergeRow:'', mergeCol:'',style:{background:'#f7f7f7',align:'center',fontFamily:'B Nazanin',borderSize:15}},   //mergeRow:(x) ,,,, mergeCol:(y)
             {x: 6, y: 1, value: '2151202', mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 6, y: 2, value: '5256636',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 6, y: 3, value: '55858412',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 6, y: 4, value: '896544',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 6, y: 5, value: '845632 ',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 6, y: 6, value: '214522',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
 
 
 
             {x: 7, y: 0, value: 'یزد',mergeRow:'', mergeCol:'',style:{background:'#f7f7f7',align:'center',fontFamily:'B Nazanin',borderSize:15}},   //mergeRow:(x) ,,,, mergeCol:(y)
             {x: 7, y: 1, value: '15554', mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 7, y: 2, value: '2256655',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 7, y: 3, value: '32265',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 7, y: 4, value: '2322565',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 7, y: 5, value: '3232265',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
             {x: 7, y: 6, value: '25520',mergeRow:'', mergeCol:'',style:{align:'center',fontFamily:'B Nazanin'}},
         ] // data
 
 
         let docx = new Docx('temp4.docx', 'outputUnitTestTmp/');
         docx.createTable(data);
         let out = docx.generate();
         if(out == false){
             fail("Don't create File of docx");
         }else{
             console.log("Create File of Docx");
             done();
         }
 
 
 
     }); */ // it  فایل مربوط به تفکیک جمعیت شهرستان
    it('Template Four', function (done) {
        var data = [
            { x: 1, y: 0, value: 'سال و استان', mergeRow: '', mergeCol: '5', style: { align: 'center', fontFamily: 'B Nazanin', topBorderSize: 17, bottomBorderSize: 17, rightBorder: 'false', leftBorderSize: 17 } },
            { x: 1, y: 1, value: '', mergeRow: '', mergeCol: '' },
            { x: 1, y: 2, value: '', mergeRow: '', mergeCol: '' },
            { x: 1, y: 3, value: '', mergeRow: '', mergeCol: '' },
            { x: 1, y: 4, value: '', mergeRow: '', mergeCol: '' },
            { x: 1, y: 5, value: 'جمع', mergeRow: '', mergeCol: '', style: { align: 'center', fontFamily: 'B Nazanin', topBorderSize: 17, rightBorderSize: 17, leftBorderSize: 17, bottomBorderSize: 17 } },
            { x: 1, y: 6, value: 'پزشک', mergeRow: '', mergeCol: '', style: { align: 'center', fontFamily: 'B Nazanin', topBorderSize: 17, rightBorderSize: 17, leftBorderSize: 17, bottomBorderSize: 17 } },
            { x: 1, y: 7, value: 'پیراپزشک', mergeRow: '', mergeCol: '', style: { align: 'center', fontFamily: 'B Nazanin', topBorderSize: 17, rightBorderSize: 17, leftBorderSize: 17, bottomBorderSize: 17 } },
            { x: 1, y: 8, value: ' ساير كاركنان ', mergeRow: '', mergeCol: '3', style: { align: 'center', borderSize: 17, fontFamily: 'B Nazanin', topBorderSize: 17, rightBorderSize: 17, leftBorder: 'false', bottomBorderSize: 17 } },
            { x: 1, y: 9, value: '  ', mergeRow: '', mergeCol: '' },
            { x: 1, y: 10, value: '  ', mergeRow: '', mergeCol: '' },
            { x: 2, y: 0, value: '1375 -----------------------------------', mergeRow: '', mergeCol: '5', style: { align: 'right', fontFamily: 'B Nazanin', topBorderSize: 17, bottomBorder: 'false', rightBorder: 'false', leftBorderSize: 17 } },
            { x: 2, y: 1, value: '', mergeRow: '', mergeCol: '' },
            { x: 2, y: 2, value: '', mergeRow: '', mergeCol: '' },
            { x: 2, y: 3, value: '', mergeRow: '', mergeCol: '' },
            { x: 2, y: 4, value: '', mergeRow: '', mergeCol: '' },
            { x: 2, y: 5, value: '269894', mergeRow: '', mergeCol: '', style: { align: 'right', fontFamily: 'B Nazanin', bottomBorder: 'false', leftBorder: 'false' } },
            { x: 2, y: 6, value: '19585', mergeRow: '', mergeCol: '', style: { align: 'right', fontFamily: 'B Nazanin', bottomBorder: 'false', leftBorder: 'false', rightBorder: 'false' } },
            { x: 2, y: 7, value: '149380', mergeRow: '', mergeCol: '', style: { align: 'right', fontFamily: 'B Nazanin', bottomBorder: 'false', leftBorder: 'false', rightBorder: 'false' } },
            { x: 2, y: 8, value: '100929', mergeRow: '', mergeCol: '3', style: { align: 'right', fontFamily: 'B Nazanin', bottomBorder: 'false', leftBorder: 'false', rightBorder: 'false' } },
            { x: 2, y: 9, value: '  ', mergeRow: '', mergeCol: '' },
            { x: 2, y: 10, value: '  ', mergeRow: '', mergeCol: '' },
            //
            //
            { x: 3, y: 0, value: '1380 -----------------------------------', mergeRow: '', mergeCol: '5', style: { align: 'right', fontFamily: 'B Nazanin', topBorder: 'false', bottomBorder: 'false', rightBorder: 'false', leftBorderSize: 17 } },
            { x: 3, y: 1, value: '', mergeRow: '', mergeCol: '' },
            { x: 3, y: 2, value: '', mergeRow: '', mergeCol: '' },
            { x: 3, y: 3, value: '', mergeRow: '', mergeCol: '' },
            { x: 3, y: 4, value: '', mergeRow: '', mergeCol: '' },
            { x: 3, y: 5, value: '295325', mergeRow: '', mergeCol: '', style: { align: 'right', fontFamily: 'B Nazanin', bottomBorder: 'false', leftBorder: 'false', topBorder: 'false' } },
            { x: 3, y: 6, value: '21175', mergeRow: '', mergeCol: '', style: { align: 'right', fontFamily: 'B Nazanin', bottomBorder: 'false', leftBorder: 'false', rightBorder: 'false', topBorder: 'false' } },
            { x: 3, y: 7, value: '152396', mergeRow: '', mergeCol: '', style: { align: 'right', fontFamily: 'B Nazanin', bottomBorder: 'false', leftBorder: 'false', rightBorder: 'false', topBorder: 'false' } },
            { x: 3, y: 8, value: '121754', mergeRow: '', mergeCol: '3', style: { align: 'right', fontFamily: 'B Nazanin', bottomBorder: 'false', leftBorder: 'false', rightBorder: 'false', topBorder: 'false' } },
            { x: 3, y: 9, value: '  ', mergeRow: '', mergeCol: '' },
            { x: 3, y: 10, value: '  ', mergeRow: '', mergeCol: '' },
            // // //
            // // //
            // // //
            { x: 4, y: 0, value: '1385 -----------------------------------', mergeRow: '', mergeCol: '5', style: { align: 'right', fontFamily: 'B Nazanin', topBorder: 'false', bottomBorder: 'false', rightBorder: 'false', leftBorderSize: 17 } },
            { x: 4, y: 1, value: '', mergeRow: '', mergeCol: '' },
            { x: 4, y: 2, value: '', mergeRow: '', mergeCol: '' },
            { x: 4, y: 3, value: '', mergeRow: '', mergeCol: '' },
            { x: 4, y: 4, value: '', mergeRow: '', mergeCol: '' },
            { x: 4, y: 5, value: '321544', mergeRow: '', mergeCol: '', style: { align: 'right', fontFamily: 'B Nazanin', bottomBorder: 'false', leftBorder: 'false', topBorder: 'false' } },
            { x: 4, y: 6, value: '29937', mergeRow: '', mergeCol: '', style: { align: 'right', fontFamily: 'B Nazanin', bottomBorder: 'false', leftBorder: 'false', rightBorder: 'false', topBorder: 'false' } },
            { x: 4, y: 7, value: '173076', mergeRow: '', mergeCol: '', style: { align: 'right', fontFamily: 'B Nazanin', bottomBorder: 'false', leftBorder: 'false', rightBorder: 'false', topBorder: 'false' } },
            { x: 4, y: 8, value: '118531', mergeRow: '', mergeCol: '3', style: { align: 'right', fontFamily: 'B Nazanin', bottomBorder: 'false', leftBorder: 'false', rightBorder: 'false', topBorder: 'false' } },
            { x: 4, y: 9, value: '  ', mergeRow: '', mergeCol: '' },
            { x: 4, y: 10, value: '  ', mergeRow: '', mergeCol: '' },
            // //
            // //
            { x: 5, y: 0, value: '1390 -----------------------------------', mergeRow: '', mergeCol: '5', style: { align: 'right', fontFamily: 'B Nazanin', topBorder: 'false', bottomBorder: 'false', rightBorder: 'false', leftBorderSize: 17 } },
            { x: 5, y: 1, value: '', mergeRow: '', mergeCol: '' },
            { x: 5, y: 2, value: '', mergeRow: '', mergeCol: '' },
            { x: 5, y: 3, value: '', mergeRow: '', mergeCol: '' },
            { x: 5, y: 4, value: '', mergeRow: '', mergeCol: '' },
            { x: 5, y: 5, value: '361627', mergeRow: '', mergeCol: '', style: { align: 'right', fontFamily: 'B Nazanin', bottomBorder: 'false', leftBorder: 'false', topBorder: 'false' } },
            { x: 5, y: 6, value: '32493', mergeRow: '', mergeCol: '', style: { align: 'right', fontFamily: 'B Nazanin', bottomBorder: 'false', leftBorder: 'false', rightBorder: 'false', topBorder: 'false' } },
            { x: 5, y: 7, value: '215950', mergeRow: '', mergeCol: '', style: { align: 'right', fontFamily: 'B Nazanin', bottomBorder: 'false', leftBorder: 'false', rightBorder: 'false', topBorder: 'false' } },
            { x: 5, y: 8, value: '113184', mergeRow: '', mergeCol: '3', style: { align: 'right', fontFamily: 'B Nazanin', bottomBorder: 'false', leftBorder: 'false', rightBorder: 'false', topBorder: 'false' } },
            { x: 5, y: 9, value: '  ', mergeRow: '', mergeCol: '' },
            { x: 5, y: 10, value: '  ', mergeRow: '', mergeCol: '' },
            // //
            // //
            { x: 6, y: 0, value: '1391 -----------------------------------', mergeRow: '', mergeCol: '5', style: { align: 'right', fontFamily: 'B Nazanin', topBorder: 'false', bottomBorder: 'false', rightBorder: 'false', leftBorderSize: 17 } },
            { x: 6, y: 1, value: '', mergeRow: '', mergeCol: '' },
            { x: 6, y: 2, value: '', mergeRow: '', mergeCol: '' },
            { x: 6, y: 3, value: '', mergeRow: '', mergeCol: '' },
            { x: 6, y: 4, value: '', mergeRow: '', mergeCol: '' },
            { x: 6, y: 5, value: '350394', mergeRow: '', mergeCol: '', style: { align: 'right', fontFamily: 'B Nazanin', bottomBorder: 'false', leftBorder: 'false', topBorder: 'false' } },
            { x: 6, y: 6, value: '34219', mergeRow: '', mergeCol: '', style: { align: 'right', fontFamily: 'B Nazanin', bottomBorder: 'false', leftBorder: 'false', rightBorder: 'false', topBorder: 'false' } },
            { x: 6, y: 7, value: '203993', mergeRow: '', mergeCol: '', style: { align: 'right', fontFamily: 'B Nazanin', bottomBorder: 'false', leftBorder: 'false', rightBorder: 'false', topBorder: 'false' } },
            { x: 6, y: 8, value: '112182', mergeRow: '', mergeCol: '3', style: { align: 'right', fontFamily: 'B Nazanin', bottomBorder: 'false', leftBorder: 'false', rightBorder: 'false', topBorder: 'false' } },
            { x: 6, y: 9, value: '  ', mergeRow: '', mergeCol: '' },
            { x: 6, y: 10, value: '  ', mergeRow: '', mergeCol: '' },
            // //
            // //
            { x: 7, y: 0, value: '1392 -----------------------------------', mergeRow: '', mergeCol: '5', style: { align: 'right', fontFamily: 'B Nazanin', topBorder: 'false', bottomBorder: 'false', rightBorder: 'false', leftBorderSize: 17 } },
            { x: 7, y: 1, value: '', mergeRow: '', mergeCol: '' },
            { x: 7, y: 2, value: '', mergeRow: '', mergeCol: '' },
            { x: 7, y: 3, value: '', mergeRow: '', mergeCol: '' },
            { x: 7, y: 4, value: '', mergeRow: '', mergeCol: '' },
            { x: 7, y: 5, value: '385667', mergeRow: '', mergeCol: '', style: { align: 'right', fontFamily: 'B Nazanin', bottomBorder: 'false', leftBorder: 'false', topBorder: 'false' } },
            { x: 7, y: 6, value: '37490', mergeRow: '', mergeCol: '', style: { align: 'right', fontFamily: 'B Nazanin', bottomBorder: 'false', leftBorder: 'false', rightBorder: 'false', topBorder: 'false' } },
            { x: 7, y: 7, value: '217603', mergeRow: '', mergeCol: '', style: { align: 'right', fontFamily: 'B Nazanin', bottomBorder: 'false', leftBorder: 'false', rightBorder: 'false', topBorder: 'false' } },
            { x: 7, y: 8, value: '130574', mergeRow: '', mergeCol: '3', style: { align: 'right', fontFamily: 'B Nazanin', bottomBorder: 'false', leftBorder: 'false', rightBorder: 'false', topBorder: 'false' } },
            { x: 7, y: 9, value: '  ', mergeRow: '', mergeCol: '' },
            { x: 7, y: 10, value: '  ', mergeRow: '', mergeCol: '' },
            // //
            // //
            // //
            { x: 8, y: 0, value: '1393 -----------------------------------', mergeRow: '', mergeCol: '5', style: { align: 'right', fontFamily: 'B Nazanin', topBorder: 'false', bottomBorder: 'false', rightBorder: 'false', leftBorderSize: 17 } },
            { x: 8, y: 1, value: '', mergeRow: '', mergeCol: '' },
            { x: 8, y: 2, value: '', mergeRow: '', mergeCol: '' },
            { x: 8, y: 3, value: '', mergeRow: '', mergeCol: '' },
            { x: 8, y: 4, value: '', mergeRow: '', mergeCol: '' },
            { x: 8, y: 5, value: '413550', mergeRow: '', mergeCol: '', style: { align: 'right', fontFamily: 'B Nazanin', bottomBorder: 'false', leftBorder: 'false', topBorder: 'false' } },
            { x: 8, y: 6, value: '42108', mergeRow: '', mergeCol: '', style: { align: 'right', fontFamily: 'B Nazanin', bottomBorder: 'false', leftBorder: 'false', rightBorder: 'false', topBorder: 'false' } },
            { x: 8, y: 7, value: '233668', mergeRow: '', mergeCol: '', style: { align: 'right', fontFamily: 'B Nazanin', bottomBorder: 'false', leftBorder: 'false', rightBorder: 'false', topBorder: 'false' } },
            { x: 8, y: 8, value: '137774', mergeRow: '', mergeCol: '3', style: { align: 'right', fontFamily: 'B Nazanin', bottomBorder: 'false', leftBorder: 'false', rightBorder: 'false', topBorder: 'false' } },
            { x: 8, y: 9, value: '  ', mergeRow: '', mergeCol: '' },
            { x: 8, y: 10, value: '  ', mergeRow: '', mergeCol: '' },
            // //
            // //
            // //
            // //
            { x: 9, y: 0, value: '1394 -----------------------------------', mergeRow: '', mergeCol: '5', style: { align: 'right', fontFamily: 'B Nazanin', topBorder: 'false', bottomBorder: 'false', rightBorder: 'false', leftBorderSize: 17 } },
            { x: 9, y: 1, value: '', mergeRow: '', mergeCol: '' },
            { x: 9, y: 2, value: '', mergeRow: '', mergeCol: '' },
            { x: 9, y: 3, value: '', mergeRow: '', mergeCol: '' },
            { x: 9, y: 4, value: '', mergeRow: '', mergeCol: '' },
            { x: 9, y: 5, value: '405910', mergeRow: '', mergeCol: '', style: { align: 'right', fontFamily: 'B Nazanin', bottomBorder: 'false', leftBorder: 'false', topBorder: 'false' } },
            { x: 9, y: 6, value: '42393', mergeRow: '', mergeCol: '', style: { align: 'right', fontFamily: 'B Nazanin', bottomBorder: 'false', leftBorder: 'false', rightBorder: 'false', topBorder: 'false' } },
            { x: 9, y: 7, value: '238869', mergeRow: '', mergeCol: '', style: { align: 'right', fontFamily: 'B Nazanin', bottomBorder: 'false', leftBorder: 'false', rightBorder: 'false', topBorder: 'false' } },
            { x: 9, y: 8, value: '124648', mergeRow: '', mergeCol: '3', style: { align: 'right', fontFamily: 'B Nazanin', bottomBorder: 'false', leftBorder: 'false', rightBorder: 'false', topBorder: 'false' } },
            { x: 9, y: 9, value: '  ', mergeRow: '', mergeCol: '' },
            { x: 9, y: 10, value: '  ', mergeRow: '', mergeCol: '' },
            // //
            // //
            { x: 10, y: 0, value: 'آذربایجان شرقی ---------------------------', mergeRow: '', mergeCol: '5', style: { align: 'right', fontFamily: 'B Nazanin', topBorder: 'false', bottomBorder: 'false', rightBorder: 'false', leftBorderSize: 17 } },
            { x: 10, y: 1, value: '', mergeRow: '', mergeCol: '' },
            { x: 10, y: 2, value: '', mergeRow: '', mergeCol: '' },
            { x: 10, y: 3, value: '', mergeRow: '', mergeCol: '' },
            { x: 10, y: 4, value: '', mergeRow: '', mergeCol: '' },
            { x: 10, y: 5, value: '22075', mergeRow: '', mergeCol: '', style: { align: 'right', fontFamily: 'B Nazanin', bottomBorder: 'false', leftBorder: 'false', topBorder: 'false' } },
            { x: 10, y: 6, value: '2262', mergeRow: '', mergeCol: '', style: { align: 'right', fontFamily: 'B Nazanin', bottomBorder: 'false', leftBorder: 'false', rightBorder: 'false', topBorder: 'false' } },
            { x: 10, y: 7, value: '13195', mergeRow: '', mergeCol: '', style: { align: 'right', fontFamily: 'B Nazanin', bottomBorder: 'false', leftBorder: 'false', rightBorder: 'false', topBorder: 'false' } },
            { x: 10, y: 8, value: '6618', mergeRow: '', mergeCol: '3', style: { align: 'right', fontFamily: 'B Nazanin', bottomBorder: 'false', leftBorder: 'false', rightBorder: 'false', topBorder: 'false' } },
            { x: 10, y: 9, value: '', mergeRow: '', mergeCol: '' },
            { x: 10, y: 10, value: '', mergeRow: '', mergeCol: '' },
            //
            // //
            // //
            { x: 11, y: 0, value: 'آذربایجان غربی ----------------------------', mergeRow: '', mergeCol: '5', style: { align: 'right', fontFamily: 'B Nazanin', topBorder: 'false', bottomBorder: 'false', rightBorder: 'false', leftBorderSize: 17 } },
            { x: 11, y: 1, value: '', mergeRow: '', mergeCol: '' },
            { x: 11, y: 2, value: '', mergeRow: '', mergeCol: '' },
            { x: 11, y: 3, value: '', mergeRow: '', mergeCol: '' },
            { x: 11, y: 4, value: '', mergeRow: '', mergeCol: '' },
            { x: 11, y: 5, value: '16483', mergeRow: '', mergeCol: '', style: { align: 'right', fontFamily: 'B Nazanin', bottomBorder: 'false', leftBorder: 'false', topBorder: 'false' } },
            { x: 11, y: 6, value: '1595', mergeRow: '', mergeCol: '', style: { align: 'right', fontFamily: 'B Nazanin', bottomBorder: 'false', leftBorder: 'false', rightBorder: 'false', topBorder: 'false' } },
            { x: 11, y: 7, value: '10924', mergeRow: '', mergeCol: '', style: { align: 'right', fontFamily: 'B Nazanin', bottomBorder: 'false', leftBorder: 'false', rightBorder: 'false', topBorder: 'false' } },
            { x: 11, y: 8, value: '3964', mergeRow: '', mergeCol: '3', style: { align: 'right', fontFamily: 'B Nazanin', bottomBorder: 'false', leftBorder: 'false', rightBorder: 'false', topBorder: 'false' } },
            { x: 11, y: 9, value: '', mergeRow: '', mergeCol: '' },
            { x: 11, y: 10, value: '', mergeRow: '', mergeCol: '' },
            // //
            // //
            // //
            { x: 12, y: 0, value: 'اردبیل -----------------------------------', mergeRow: '', mergeCol: '5', style: { align: 'right', fontFamily: 'B Nazanin', topBorder: 'false', bottomBorder: 'false', rightBorder: 'false', leftBorderSize: 17 } },
            { x: 12, y: 1, value: '', mergeRow: '', mergeCol: '' },
            { x: 12, y: 2, value: '', mergeRow: '', mergeCol: '' },
            { x: 12, y: 3, value: '', mergeRow: '', mergeCol: '' },
            { x: 12, y: 4, value: '', mergeRow: '', mergeCol: '' },
            { x: 12, y: 5, value: '6755', mergeRow: '', mergeCol: '', style: { align: 'right', fontFamily: 'B Nazanin', bottomBorder: 'false', leftBorder: 'false', topBorder: 'false' } },
            { x: 12, y: 6, value: '487', mergeRow: '', mergeCol: '', style: { align: 'right', fontFamily: 'B Nazanin', bottomBorder: 'false', leftBorder: 'false', rightBorder: 'false', topBorder: 'false' } },
            { x: 12, y: 7, value: '4565', mergeRow: '', mergeCol: '', style: { align: 'right', fontFamily: 'B Nazanin', bottomBorder: 'false', leftBorder: 'false', rightBorder: 'false', topBorder: 'false' } },
            { x: 12, y: 8, value: '1703', mergeRow: '', mergeCol: '3', style: { align: 'right', fontFamily: 'B Nazanin', bottomBorder: 'false', leftBorder: 'false', rightBorder: 'false', topBorder: 'false' } },
            { x: 12, y: 9, value: '', mergeRow: '', mergeCol: '' },
            { x: 12, y: 10, value: '', mergeRow: '', mergeCol: '' },
            // //
            // //
            // //
            { x: 13, y: 0, value: 'اصفهان ----------------------------------', mergeRow: '', mergeCol: '5', style: { align: 'right', fontFamily: 'B Nazanin', topBorder: 'false', bottomBorder: 'false', rightBorder: 'false', leftBorderSize: 17 } },
            { x: 13, y: 1, value: '', mergeRow: '', mergeCol: '' },
            { x: 13, y: 2, value: '', mergeRow: '', mergeCol: '' },
            { x: 13, y: 3, value: '', mergeRow: '', mergeCol: '' },
            { x: 13, y: 4, value: '', mergeRow: '', mergeCol: '' },
            { x: 13, y: 5, value: '28384', mergeRow: '', mergeCol: '', style: { align: 'right', fontFamily: 'B Nazanin', bottomBorder: 'false', leftBorder: 'false', topBorder: 'false' } },
            { x: 13, y: 6, value: '3218', mergeRow: '', mergeCol: '', style: { align: 'right', fontFamily: 'B Nazanin', bottomBorder: 'false', leftBorder: 'false', rightBorder: 'false', topBorder: 'false' } },
            { x: 13, y: 7, value: '17026', mergeRow: '', mergeCol: '', style: { align: 'right', fontFamily: 'B Nazanin', bottomBorder: 'false', leftBorder: 'false', rightBorder: 'false', topBorder: 'false' } },
            { x: 13, y: 8, value: '17026', mergeRow: '', mergeCol: '3', style: { align: 'right', fontFamily: 'B Nazanin', bottomBorder: 'false', leftBorder: 'false', rightBorder: 'false', topBorder: 'false' } },
            { x: 13, y: 9, value: '', mergeRow: '', mergeCol: '' },
            { x: 13, y: 10, value: '', mergeRow: '', mergeCol: '' },
            // //
            // //
            // //
            // //
            { x: 14, y: 0, value: 'البرز -----------------------------------', mergeRow: '', mergeCol: '5', style: { align: 'right', fontFamily: 'B Nazanin', topBorder: 'false', bottomBorder: 'false', rightBorder: 'false', leftBorderSize: 17 } },
            { x: 14, y: 1, value: '', mergeRow: '', mergeCol: '' },
            { x: 14, y: 2, value: '', mergeRow: '', mergeCol: '' },
            { x: 14, y: 3, value: '', mergeRow: '', mergeCol: '' },
            { x: 14, y: 4, value: '', mergeRow: '', mergeCol: '' },
            { x: 14, y: 5, value: '5961', mergeRow: '', mergeCol: '', style: { align: 'right', fontFamily: 'B Nazanin', bottomBorder: 'false', leftBorder: 'false', topBorder: 'false' } },
            { x: 14, y: 6, value: '911', mergeRow: '', mergeCol: '', style: { align: 'right', fontFamily: 'B Nazanin', bottomBorder: 'false', leftBorder: 'false', rightBorder: 'false', topBorder: 'false' } },
            { x: 14, y: 7, value: '2712', mergeRow: '', mergeCol: '', style: { align: 'right', fontFamily: 'B Nazanin', bottomBorder: 'false', leftBorder: 'false', rightBorder: 'false', topBorder: 'false' } },
            { x: 14, y: 8, value: '2338', mergeRow: '', mergeCol: '3', style: { align: 'right', fontFamily: 'B Nazanin', bottomBorder: 'false', leftBorder: 'false', rightBorder: 'false', topBorder: 'false' } },
            { x: 14, y: 9, value: '', mergeRow: '', mergeCol: '' },
            { x: 14, y: 10, value: '', mergeRow: '', mergeCol: '' },
            // //
            // //
            // //
            // //
            // //
            { x: 15, y: 0, value: 'ایلام------------------------------------', mergeRow: '', mergeCol: '5', style: { align: 'right', fontFamily: 'B Nazanin', topBorder: 'false', bottomBorder: 'false', rightBorder: 'false', leftBorderSize: 17 } },
            { x: 15, y: 1, value: '', mergeRow: '', mergeCol: '' },
            { x: 15, y: 2, value: '', mergeRow: '', mergeCol: '' },
            { x: 15, y: 3, value: '', mergeRow: '', mergeCol: '' },
            { x: 15, y: 4, value: '', mergeRow: '', mergeCol: '' },
            { x: 15, y: 5, value: '4545', mergeRow: '', mergeCol: '', style: { align: 'right', fontFamily: 'B Nazanin', bottomBorder: 'false', leftBorder: 'false', topBorder: 'false' } },
            { x: 15, y: 6, value: '388', mergeRow: '', mergeCol: '', style: { align: 'right', fontFamily: 'B Nazanin', bottomBorder: 'false', leftBorder: 'false', rightBorder: 'false', topBorder: 'false' } },
            { x: 15, y: 7, value: '3100', mergeRow: '', mergeCol: '', style: { align: 'right', fontFamily: 'B Nazanin', bottomBorder: 'false', leftBorder: 'false', rightBorder: 'false', topBorder: 'false' } },
            { x: 15, y: 8, value: '1057', mergeRow: '', mergeCol: '3', style: { align: 'right', fontFamily: 'B Nazanin', bottomBorder: 'false', leftBorder: 'false', rightBorder: 'false', topBorder: 'false' } },
            { x: 15, y: 9, value: '', mergeRow: '', mergeCol: '' },
            { x: 15, y: 10, value: '', mergeRow: '', mergeCol: '' },
            // //
            // //
            // //
            { x: 16, y: 0, value: 'بوشهر -----------------------------------', mergeRow: '', mergeCol: '5', style: { align: 'right', fontFamily: 'B Nazanin', topBorder: 'false', bottomBorder: 'false', rightBorder: 'false', leftBorderSize: 17 } },
            { x: 16, y: 1, value: '', mergeRow: '', mergeCol: '' },
            { x: 16, y: 2, value: '', mergeRow: '', mergeCol: '' },
            { x: 16, y: 3, value: '', mergeRow: '', mergeCol: '' },
            { x: 16, y: 4, value: '', mergeRow: '', mergeCol: '' },
            { x: 16, y: 5, value: '6877', mergeRow: '', mergeCol: '', style: { align: 'right', fontFamily: 'B Nazanin', bottomBorder: 'false', leftBorder: 'false', topBorder: 'false' } },
            { x: 16, y: 6, value: '704', mergeRow: '', mergeCol: '', style: { align: 'right', fontFamily: 'B Nazanin', bottomBorder: 'false', leftBorder: 'false', rightBorder: 'false', topBorder: 'false' } },
            { x: 16, y: 7, value: '3916', mergeRow: '', mergeCol: '', style: { align: 'right', fontFamily: 'B Nazanin', bottomBorder: 'false', leftBorder: 'false', rightBorder: 'false', topBorder: 'false' } },
            { x: 16, y: 8, value: '2257', mergeRow: '', mergeCol: '3', style: { align: 'right', fontFamily: 'B Nazanin', bottomBorder: 'false', leftBorder: 'false', rightBorder: 'false', topBorder: 'false' } },
            { x: 16, y: 9, value: '', mergeRow: '', mergeCol: '' },
            { x: 16, y: 10, value: '', mergeRow: '', mergeCol: '' },
            //
            // //
            // //
            // //
            // //
            { x: 17, y: 0, value: 'تهران -----------------------------------', mergeRow: '', mergeCol: '5', style: { align: 'right', fontFamily: 'B Nazanin', topBorder: 'false', bottomBorder: 'false', rightBorder: 'false', leftBorderSize: 17 } },
            { x: 17, y: 1, value: '', mergeRow: '', mergeCol: '' },
            { x: 17, y: 2, value: '', mergeRow: '', mergeCol: '' },
            { x: 17, y: 2, value: '', mergeRow: '', mergeCol: '' },
            { x: 17, y: 3, value: '', mergeRow: '', mergeCol: '' },
            { x: 17, y: 4, value: '', mergeRow: '', mergeCol: '' },
            { x: 17, y: 5, value: '47440', mergeRow: '', mergeCol: '', style: { align: 'right', fontFamily: 'B Nazanin', bottomBorder: 'false', leftBorder: 'false', topBorder: 'false' } },
            { x: 17, y: 6, value: '6313', mergeRow: '', mergeCol: '', style: { align: 'right', fontFamily: 'B Nazanin', bottomBorder: 'false', leftBorder: 'false', rightBorder: 'false', topBorder: 'false' } },
            { x: 17, y: 7, value: '22892', mergeRow: '', mergeCol: '', style: { align: 'right', fontFamily: 'B Nazanin', bottomBorder: 'false', leftBorder: 'false', rightBorder: 'false', topBorder: 'false' } },
            { x: 17, y: 8, value: '18235', mergeRow: '', mergeCol: '3', style: { align: 'right', fontFamily: 'B Nazanin', bottomBorder: 'false', leftBorder: 'false', rightBorder: 'false', topBorder: 'false' } },
            { x: 17, y: 9, value: '', mergeRow: '', mergeCol: '' },
            { x: 17, y: 10, value: '', mergeRow: '', mergeCol: '' },
            // //
            // //
            // //
            { x: 18, y: 0, value: 'چهارمحال و بختیاری ----------------------', mergeRow: '', mergeCol: '5', style: { align: 'right', fontFamily: 'B Nazanin', topBorder: 'false', bottomBorder: 'false', rightBorder: 'false', leftBorderSize: 17 } },
            { x: 18, y: 1, value: '', mergeRow: '', mergeCol: '' },
            { x: 18, y: 2, value: '', mergeRow: '', mergeCol: '' },
            { x: 18, y: 3, value: '', mergeRow: '', mergeCol: '' },
            { x: 18, y: 4, value: '', mergeRow: '', mergeCol: '' },
            { x: 18, y: 5, value: '7175', mergeRow: '', mergeCol: '', style: { align: 'right', fontFamily: 'B Nazanin', bottomBorder: 'false', leftBorder: 'false', topBorder: 'false' } },
            { x: 18, y: 6, value: '752', mergeRow: '', mergeCol: '', style: { align: 'right', fontFamily: 'B Nazanin', bottomBorder: 'false', leftBorder: 'false', rightBorder: 'false', topBorder: 'false' } },
            { x: 18, y: 7, value: '4388', mergeRow: '', mergeCol: '', style: { align: 'right', fontFamily: 'B Nazanin', bottomBorder: 'false', leftBorder: 'false', rightBorder: 'false', topBorder: 'false' } },
            { x: 18, y: 8, value: '2035', mergeRow: '', mergeCol: '3', style: { align: 'right', fontFamily: 'B Nazanin', bottomBorder: 'false', leftBorder: 'false', rightBorder: 'false', topBorder: 'false' } },
            { x: 18, y: 9, value: '', mergeRow: '', mergeCol: '' },
            { x: 18, y: 10, value: '', mergeRow: '', mergeCol: '' },
            // //
            // //
            // //
            { x: 19, y: 0, value: 'خراسان جنوبی ----------------------------', mergeRow: '', mergeCol: '5', style: { align: 'right', fontFamily: 'B Nazanin', topBorder: 'false', bottomBorder: 'false', rightBorder: 'false', leftBorderSize: 17 } },
            { x: 19, y: 1, value: '', mergeRow: '', mergeCol: '' },
            { x: 19, y: 2, value: '', mergeRow: '', mergeCol: '' },
            { x: 19, y: 3, value: '', mergeRow: '', mergeCol: '' },
            { x: 19, y: 4, value: '', mergeRow: '', mergeCol: '' },
            { x: 19, y: 5, value: '5167', mergeRow: '', mergeCol: '', style: { align: 'right', fontFamily: 'B Nazanin', bottomBorder: 'false', leftBorder: 'false', topBorder: 'false' } },
            { x: 19, y: 6, value: '542', mergeRow: '', mergeCol: '', style: { align: 'right', fontFamily: 'B Nazanin', bottomBorder: 'false', leftBorder: 'false', rightBorder: 'false', topBorder: 'false' } },
            { x: 19, y: 7, value: '3275', mergeRow: '', mergeCol: '', style: { align: 'right', fontFamily: 'B Nazanin', bottomBorder: 'false', leftBorder: 'false', rightBorder: 'false', topBorder: 'false' } },
            { x: 19, y: 8, value: '1350', mergeRow: '', mergeCol: '3', style: { align: 'right', fontFamily: 'B Nazanin', bottomBorder: 'false', leftBorder: 'false', rightBorder: 'false', topBorder: 'false' } },
            { x: 19, y: 9, value: '', mergeRow: '', mergeCol: '' },
            { x: 19, y: 10, value: '', mergeRow: '', mergeCol: '' },
            // //
            // //
            // //
            { x: 20, y: 0, value: 'خراسان رضوی -----------------------------', mergeRow: '', mergeCol: '5', style: { align: 'right', fontFamily: 'B Nazanin', topBorder: 'false', bottomBorder: 'false', rightBorder: 'false', leftBorderSize: 17 } },
            { x: 20, y: 1, value: '', mergeRow: '', mergeCol: '' },
            { x: 20, y: 2, value: '', mergeRow: '', mergeCol: '' },
            { x: 20, y: 3, value: '', mergeRow: '', mergeCol: '' },
            { x: 20, y: 4, value: '', mergeRow: '', mergeCol: '' },
            { x: 20, y: 5, value: '30302', mergeRow: '', mergeCol: '', style: { align: 'right', fontFamily: 'B Nazanin', bottomBorder: 'false', leftBorder: 'false', topBorder: 'false' } },
            { x: 20, y: 6, value: '3266', mergeRow: '', mergeCol: '', style: { align: 'right', fontFamily: 'B Nazanin', bottomBorder: 'false', leftBorder: 'false', rightBorder: 'false', topBorder: 'false' } },
            { x: 20, y: 7, value: '18483', mergeRow: '', mergeCol: '', style: { align: 'right', fontFamily: 'B Nazanin', bottomBorder: 'false', leftBorder: 'false', rightBorder: 'false', topBorder: 'false' } },
            { x: 20, y: 8, value: '8553', mergeRow: '', mergeCol: '3', style: { align: 'right', fontFamily: 'B Nazanin', bottomBorder: 'false', leftBorder: 'false', rightBorder: 'false', topBorder: 'false' } },
            { x: 20, y: 9, value: '', mergeRow: '', mergeCol: '' },
            { x: 20, y: 10, value: '', mergeRow: '', mergeCol: '' },
            // //
            // //
            // //
            { x: 21, y: 0, value: 'خراسان شمالی  ---------------------------', mergeRow: '', mergeCol: '5', style: { align: 'right', fontFamily: 'B Nazanin', topBorder: 'false', bottomBorder: 'false', rightBorder: 'false', leftBorderSize: 17 } },
            { x: 21, y: 1, value: '', mergeRow: '', mergeCol: '' },
            { x: 21, y: 2, value: '', mergeRow: '', mergeCol: '' },
            { x: 21, y: 3, value: '', mergeRow: '', mergeCol: '' },
            { x: 21, y: 4, value: '', mergeRow: '', mergeCol: '' },
            { x: 21, y: 5, value: '4913', mergeRow: '', mergeCol: '', style: { align: 'right', fontFamily: 'B Nazanin', bottomBorder: 'false', leftBorder: 'false', topBorder: 'false' } },
            { x: 21, y: 6, value: '499', mergeRow: '', mergeCol: '', style: { align: 'right', fontFamily: 'B Nazanin', bottomBorder: 'false', leftBorder: 'false', rightBorder: 'false', topBorder: 'false' } },
            { x: 21, y: 7, value: '3205', mergeRow: '', mergeCol: '', style: { align: 'right', fontFamily: 'B Nazanin', bottomBorder: 'false', leftBorder: 'false', rightBorder: 'false', topBorder: 'false' } },
            { x: 21, y: 8, value: '1209', mergeRow: '', mergeCol: '3', style: { align: 'right', fontFamily: 'B Nazanin', bottomBorder: 'false', leftBorder: 'false', rightBorder: 'false', topBorder: 'false' } },
            { x: 21, y: 9, value: '', mergeRow: '', mergeCol: '' },
            { x: 21, y: 10, value: '', mergeRow: '', mergeCol: '' },
            // //
            // //
            // //
            { x: 22, y: 0, value: 'خوزستان  -----------------------------', mergeRow: '', mergeCol: '5', style: { align: 'right', fontFamily: 'B Nazanin', topBorder: 'false', bottomBorder: 'false', rightBorder: 'false', leftBorderSize: 17 } },
            { x: 22, y: 1, value: '', mergeRow: '', mergeCol: '' },
            { x: 22, y: 2, value: '', mergeRow: '', mergeCol: '' },
            { x: 22, y: 3, value: '', mergeRow: '', mergeCol: '' },
            { x: 22, y: 4, value: '', mergeRow: '', mergeCol: '' },
            { x: 22, y: 5, value: '7180', mergeRow: '', mergeCol: '', style: { align: 'right', fontFamily: 'B Nazanin', bottomBorder: 'false', leftBorder: 'false', topBorder: 'false' } },
            { x: 22, y: 6, value: '715', mergeRow: '', mergeCol: '', style: { align: 'right', fontFamily: 'B Nazanin', bottomBorder: 'false', leftBorder: 'false', rightBorder: 'false', topBorder: 'false' } },
            { x: 22, y: 7, value: '4449', mergeRow: '', mergeCol: '', style: { align: 'right', fontFamily: 'B Nazanin', bottomBorder: 'false', leftBorder: 'false', rightBorder: 'false', topBorder: 'false' } },
            { x: 22, y: 8, value: '2016', mergeRow: '', mergeCol: '3', style: { align: 'right', fontFamily: 'B Nazanin', bottomBorder: 'false', leftBorder: 'false', rightBorder: 'false', topBorder: 'false' } },
            { x: 22, y: 9, value: '', mergeRow: '', mergeCol: '' },
            { x: 22, y: 10, value: '', mergeRow: '', mergeCol: '' },
            // //
            // //
            // //
            { x: 23, y: 0, value: 'زنجان -------------------------------', mergeRow: '', mergeCol: '5', style: { align: 'right', fontFamily: 'B Nazanin', topBorder: 'false', bottomBorder: 'false', rightBorder: 'false', leftBorderSize: 17 } },
            { x: 23, y: 1, value: '', mergeRow: '', mergeCol: '' },
            { x: 23, y: 2, value: '', mergeRow: '', mergeCol: '' },
            { x: 23, y: 3, value: '', mergeRow: '', mergeCol: '' },
            { x: 23, y: 4, value: '', mergeRow: '', mergeCol: '' },
            { x: 23, y: 5, value: '7007', mergeRow: '', mergeCol: '', style: { align: 'right', fontFamily: 'B Nazanin', bottomBorder: 'false', leftBorder: 'false', topBorder: 'false' } },
            { x: 23, y: 6, value: '751', mergeRow: '', mergeCol: '', style: { align: 'right', fontFamily: 'B Nazanin', bottomBorder: 'false', leftBorder: 'false', rightBorder: 'false', topBorder: 'false' } },
            { x: 23, y: 7, value: '4135', mergeRow: '', mergeCol: '', style: { align: 'right', fontFamily: 'B Nazanin', bottomBorder: 'false', leftBorder: 'false', rightBorder: 'false', topBorder: 'false' } },
            { x: 23, y: 8, value: '2121', mergeRow: '', mergeCol: '3', style: { align: 'right', fontFamily: 'B Nazanin', bottomBorder: 'false', leftBorder: 'false', rightBorder: 'false', topBorder: 'false' } },
            { x: 23, y: 9, value: '', mergeRow: '', mergeCol: '' },
            { x: 23, y: 10, value: '', mergeRow: '', mergeCol: '' },
            // //
            // //
            // //
            { x: 24, y: 0, value: 'سمنان -----------------------------------', mergeRow: '', mergeCol: '5', style: { align: 'right', fontFamily: 'B Nazanin', topBorder: 'false', bottomBorder: 'false', rightBorder: 'false', leftBorderSize: 17 } },
            { x: 24, y: 1, value: '', mergeRow: '', mergeCol: '' },
            { x: 24, y: 2, value: '', mergeRow: '', mergeCol: '' },
            { x: 24, y: 3, value: '', mergeRow: '', mergeCol: '' },
            { x: 24, y: 4, value: '', mergeRow: '', mergeCol: '' },
            { x: 24, y: 5, value: '4425', mergeRow: '', mergeCol: '', style: { align: 'right', fontFamily: 'B Nazanin', bottomBorder: 'false', leftBorder: 'false', topBorder: 'false' } },
            { x: 24, y: 6, value: '648', mergeRow: '', mergeCol: '', style: { align: 'right', fontFamily: 'B Nazanin', bottomBorder: 'false', leftBorder: 'false', rightBorder: 'false', topBorder: 'false' } },
            { x: 24, y: 7, value: '1778', mergeRow: '', mergeCol: '', style: { align: 'right', fontFamily: 'B Nazanin', bottomBorder: 'false', leftBorder: 'false', rightBorder: 'false', topBorder: 'false' } },
            { x: 24, y: 8, value: '1999', mergeRow: '', mergeCol: '3', style: { align: 'right', fontFamily: 'B Nazanin', bottomBorder: 'false', leftBorder: 'false', rightBorder: 'false', topBorder: 'false' } },
            { x: 24, y: 9, value: '', mergeRow: '', mergeCol: '' },
            { x: 24, y: 10, value: '', mergeRow: '', mergeCol: '' },
            //
            // //
            // //
            // //
            { x: 25, y: 0, value: 'سیستان و بلوچستان -----------------------', mergeRow: '', mergeCol: '5', style: { align: 'right', fontFamily: 'B Nazanin', topBorder: 'false', bottomBorder: 'false', rightBorder: 'false', leftBorderSize: 17 } },
            { x: 25, y: 1, value: '', mergeRow: '', mergeCol: '' },
            { x: 25, y: 2, value: '', mergeRow: '', mergeCol: '' },
            { x: 25, y: 3, value: '', mergeRow: '', mergeCol: '' },
            { x: 25, y: 4, value: '', mergeRow: '', mergeCol: '' },
            { x: 25, y: 5, value: '14437', mergeRow: '', mergeCol: '', style: { align: 'right', fontFamily: 'B Nazanin', bottomBorder: 'false', leftBorder: 'false', topBorder: 'false' } },
            { x: 25, y: 6, value: '1204', mergeRow: '', mergeCol: '', style: { align: 'right', fontFamily: 'B Nazanin', bottomBorder: 'false', leftBorder: 'false', rightBorder: 'false', topBorder: 'false' } },
            { x: 25, y: 7, value: '8482', mergeRow: '', mergeCol: '', style: { align: 'right', fontFamily: 'B Nazanin', bottomBorder: 'false', leftBorder: 'false', rightBorder: 'false', topBorder: 'false' } },
            { x: 25, y: 8, value: '4751', mergeRow: '', mergeCol: '3', style: { align: 'right', fontFamily: 'B Nazanin', bottomBorder: 'false', leftBorder: 'false', rightBorder: 'false', topBorder: 'false' } },
            { x: 25, y: 9, value: '', mergeRow: '', mergeCol: '' },
            { x: 25, y: 10, value: '', mergeRow: '', mergeCol: '' },
            // //
            // //
            // //
            { x: 26, y: 0, value: 'فارس -----------------------------------', mergeRow: '', mergeCol: '5', style: { align: 'right', fontFamily: 'B Nazanin', topBorder: 'false', bottomBorder: 'false', rightBorder: 'false', leftBorderSize: 17 } },
            { x: 26, y: 1, value: '', mergeRow: '', mergeCol: '' },
            { x: 26, y: 2, value: '', mergeRow: '', mergeCol: '' },
            { x: 26, y: 3, value: '', mergeRow: '', mergeCol: '' },
            { x: 26, y: 4, value: '', mergeRow: '', mergeCol: '' },
            { x: 26, y: 5, value: '32250', mergeRow: '', mergeCol: '', style: { align: 'right', fontFamily: 'B Nazanin', bottomBorder: 'false', leftBorder: 'false', topBorder: 'false' } },
            { x: 26, y: 6, value: '3026', mergeRow: '', mergeCol: '', style: { align: 'right', fontFamily: 'B Nazanin', bottomBorder: 'false', leftBorder: 'false', rightBorder: 'false', topBorder: 'false' } },
            { x: 26, y: 7, value: '19621', mergeRow: '', mergeCol: '', style: { align: 'right', fontFamily: 'B Nazanin', bottomBorder: 'false', leftBorder: 'false', rightBorder: 'false', topBorder: 'false' } },
            { x: 26, y: 8, value: '9603', mergeRow: '', mergeCol: '3', style: { align: 'right', fontFamily: 'B Nazanin', bottomBorder: 'false', leftBorder: 'false', rightBorder: 'false', topBorder: 'false' } },
            { x: 26, y: 9, value: '', mergeRow: '', mergeCol: '' },
            { x: 26, y: 10, value: '', mergeRow: '', mergeCol: '' },
            // //
            // //
            // //
            { x: 27, y: 0, value: 'قزوین -----------------------------------', mergeRow: '', mergeCol: '5', style: { align: 'right', fontFamily: 'B Nazanin', topBorder: 'false', bottomBorder: 'false', rightBorder: 'false', leftBorderSize: 17 } },
            { x: 27, y: 1, value: '', mergeRow: '', mergeCol: '' },
            { x: 27, y: 2, value: '', mergeRow: '', mergeCol: '' },
            { x: 27, y: 3, value: '', mergeRow: '', mergeCol: '' },
            { x: 27, y: 4, value: '', mergeRow: '', mergeCol: '' },
            { x: 27, y: 5, value: '6682', mergeRow: '', mergeCol: '', style: { align: 'right', fontFamily: 'B Nazanin', bottomBorder: 'false', leftBorder: 'false', topBorder: 'false' } },
            { x: 27, y: 6, value: '518', mergeRow: '', mergeCol: '', style: { align: 'right', fontFamily: 'B Nazanin', bottomBorder: 'false', leftBorder: 'false', rightBorder: 'false', topBorder: 'false' } },
            { x: 27, y: 7, value: '3899', mergeRow: '', mergeCol: '', style: { align: 'right', fontFamily: 'B Nazanin', bottomBorder: 'false', leftBorder: 'false', rightBorder: 'false', topBorder: 'false' } },
            { x: 27, y: 8, value: '2265', mergeRow: '', mergeCol: '3', style: { align: 'right', fontFamily: 'B Nazanin', bottomBorder: 'false', leftBorder: 'false', rightBorder: 'false', topBorder: 'false' } },
            { x: 27, y: 9, value: '', mergeRow: '', mergeCol: '' },
            { x: 27, y: 10, value: '', mergeRow: '', mergeCol: '' },
            // //
            // //
            // //
            { x: 28, y: 0, value: 'قم -----------------------------------', mergeRow: '', mergeCol: '5', style: { align: 'right', fontFamily: 'B Nazanin', topBorder: 'false', bottomBorder: 'false', rightBorder: 'false', leftBorderSize: 17 } },
            { x: 28, y: 1, value: '', mergeRow: '', mergeCol: '' },
            { x: 28, y: 2, value: '', mergeRow: '', mergeCol: '' },
            { x: 28, y: 3, value: '', mergeRow: '', mergeCol: '' },
            { x: 28, y: 4, value: '', mergeRow: '', mergeCol: '' },
            { x: 28, y: 5, value: '4806', mergeRow: '', mergeCol: '', style: { align: 'right', fontFamily: 'B Nazanin', bottomBorder: 'false', leftBorder: 'false', topBorder: 'false' } },
            { x: 28, y: 6, value: '555', mergeRow: '', mergeCol: '', style: { align: 'right', fontFamily: 'B Nazanin', bottomBorder: 'false', leftBorder: 'false', rightBorder: 'false', topBorder: 'false' } },
            { x: 28, y: 7, value: '2717', mergeRow: '', mergeCol: '', style: { align: 'right', fontFamily: 'B Nazanin', bottomBorder: 'false', leftBorder: 'false', rightBorder: 'false', topBorder: 'false' } },
            { x: 28, y: 8, value: '1534', mergeRow: '', mergeCol: '3', style: { align: 'right', fontFamily: 'B Nazanin', bottomBorder: 'false', leftBorder: 'false', rightBorder: 'false', topBorder: 'false' } },
            { x: 28, y: 9, value: '', mergeRow: '', mergeCol: '' },
            { x: 28, y: 10, value: '', mergeRow: '', mergeCol: '' },
            // //
            // //
            // //
            // //
            // //
            { x: 29, y: 0, value: 'کردستان --------------------------------', mergeRow: '', mergeCol: '5', style: { align: 'right', fontFamily: 'B Nazanin', topBorder: 'false', bottomBorder: 'false', rightBorder: 'false', leftBorderSize: 17 } },
            { x: 29, y: 1, value: '', mergeRow: '', mergeCol: '' },
            { x: 29, y: 2, value: '', mergeRow: '', mergeCol: '' },
            { x: 29, y: 3, value: '', mergeRow: '', mergeCol: '' },
            { x: 29, y: 4, value: '', mergeRow: '', mergeCol: '' },
            { x: 29, y: 5, value: '9003', mergeRow: '', mergeCol: '', style: { align: 'right', fontFamily: 'B Nazanin', bottomBorder: 'false', leftBorder: 'false', topBorder: 'false' } },
            { x: 29, y: 6, value: '810', mergeRow: '', mergeCol: '', style: { align: 'right', fontFamily: 'B Nazanin', bottomBorder: 'false', leftBorder: 'false', rightBorder: 'false', topBorder: 'false' } },
            { x: 29, y: 7, value: '6206', mergeRow: '', mergeCol: '', style: { align: 'right', fontFamily: 'B Nazanin', bottomBorder: 'false', leftBorder: 'false', rightBorder: 'false', topBorder: 'false' } },
            { x: 29, y: 8, value: '1987', mergeRow: '', mergeCol: '3', style: { align: 'right', fontFamily: 'B Nazanin', bottomBorder: 'false', leftBorder: 'false', rightBorder: 'false', topBorder: 'false' } },
            { x: 29, y: 9, value: '', mergeRow: '', mergeCol: '' },
            { x: 29, y: 10, value: '', mergeRow: '', mergeCol: '' },
            //
            // //
            // //
            // //
            // //
            { x: 30, y: 0, value: 'کرمان -----------------------------------', mergeRow: '', mergeCol: '5', style: { align: 'right', fontFamily: 'B Nazanin', topBorder: 'false', bottomBorder: 'false', rightBorder: 'false', leftBorderSize: 17 } },
            { x: 30, y: 1, value: '', mergeRow: '', mergeCol: '' },
            { x: 30, y: 2, value: '', mergeRow: '', mergeCol: '' },
            { x: 30, y: 3, value: '', mergeRow: '', mergeCol: '' },
            { x: 30, y: 4, value: '', mergeRow: '', mergeCol: '' },
            { x: 30, y: 5, value: '18318', mergeRow: '', mergeCol: '', style: { align: 'right', fontFamily: 'B Nazanin', bottomBorder: 'false', leftBorder: 'false', topBorder: 'false' } },
            { x: 30, y: 6, value: '1882', mergeRow: '', mergeCol: '', style: { align: 'right', fontFamily: 'B Nazanin', bottomBorder: 'false', leftBorder: 'false', rightBorder: 'false', topBorder: 'false' } },
            { x: 30, y: 7, value: '11091', mergeRow: '', mergeCol: '', style: { align: 'right', fontFamily: 'B Nazanin', bottomBorder: 'false', leftBorder: 'false', rightBorder: 'false', topBorder: 'false' } },
            { x: 30, y: 8, value: '5345', mergeRow: '', mergeCol: '3', style: { align: 'right', fontFamily: 'B Nazanin', bottomBorder: 'false', leftBorder: 'false', rightBorder: 'false', topBorder: 'false' } },
            { x: 30, y: 9, value: '', mergeRow: '', mergeCol: '' },
            { x: 30, y: 10, value: '', mergeRow: '', mergeCol: '' },
            // //
            // //
            // //
            { x: 31, y: 0, value: 'کرمانشاه -----------------------------------', mergeRow: '', mergeCol: '5', style: { align: 'right', fontFamily: 'B Nazanin', topBorder: 'false', bottomBorder: 'false', rightBorder: 'false', leftBorderSize: 17 } },
            { x: 31, y: 1, value: '', mergeRow: '', mergeCol: '' },
            { x: 31, y: 2, value: '', mergeRow: '', mergeCol: '' },
            { x: 31, y: 3, value: '', mergeRow: '', mergeCol: '' },
            { x: 31, y: 4, value: '', mergeRow: '', mergeCol: '' },
            { x: 31, y: 5, value: '12228', mergeRow: '', mergeCol: '', style: { align: 'right', fontFamily: 'B Nazanin', bottomBorder: 'false', leftBorder: 'false', topBorder: 'false' } },
            { x: 31, y: 6, value: '1213', mergeRow: '', mergeCol: '', style: { align: 'right', fontFamily: 'B Nazanin', bottomBorder: 'false', leftBorder: 'false', rightBorder: 'false', topBorder: 'false' } },
            { x: 31, y: 7, value: '7906', mergeRow: '', mergeCol: '', style: { align: 'right', fontFamily: 'B Nazanin', bottomBorder: 'false', leftBorder: 'false', rightBorder: 'false', topBorder: 'false' } },
            { x: 31, y: 8, value: '3109', mergeRow: '', mergeCol: '3', style: { align: 'right', fontFamily: 'B Nazanin', bottomBorder: 'false', leftBorder: 'false', rightBorder: 'false', topBorder: 'false' } },
            { x: 31, y: 9, value: '', mergeRow: '', mergeCol: '' },
            { x: 31, y: 10, value: '', mergeRow: '', mergeCol: '' },
            // //
            // //
            // //
            { x: 32, y: 0, value: 'كهگيلويه و بويراحمد ------------------------', mergeRow: '', mergeCol: '5', style: { align: 'right', fontFamily: 'B Nazanin', topBorder: 'false', bottomBorder: 'false', rightBorder: 'false', leftBorderSize: 17 } },
            { x: 32, y: 1, value: '', mergeRow: '', mergeCol: '' },
            { x: 32, y: 2, value: '', mergeRow: '', mergeCol: '' },
            { x: 32, y: 3, value: '', mergeRow: '', mergeCol: '' },
            { x: 32, y: 4, value: '', mergeRow: '', mergeCol: '' },
            { x: 32, y: 5, value: '5345', mergeRow: '', mergeCol: '', style: { align: 'right', fontFamily: 'B Nazanin', bottomBorder: 'false', leftBorder: 'false', topBorder: 'false' } },
            { x: 32, y: 6, value: '0', mergeRow: '', mergeCol: '', style: { align: 'right', fontFamily: 'B Nazanin', bottomBorder: 'false', leftBorder: 'false', rightBorder: 'false', topBorder: 'false' } },
            { x: 32, y: 7, value: '0', mergeRow: '', mergeCol: '', style: { align: 'right', fontFamily: 'B Nazanin', bottomBorder: 'false', leftBorder: 'false', rightBorder: 'false', topBorder: 'false' } },
            { x: 32, y: 8, value: '5345', mergeRow: '', mergeCol: '3', style: { align: 'right', fontFamily: 'B Nazanin', bottomBorder: 'false', leftBorder: 'false', rightBorder: 'false', topBorder: 'false' } },
            { x: 32, y: 9, value: '', mergeRow: '', mergeCol: '' },
            { x: 32, y: 10, value: '', mergeRow: '', mergeCol: '' },
            // //
            // //
            // //
            { x: 33, y: 0, value: 'گلستان -----------------------------------', mergeRow: '', mergeCol: '5', style: { align: 'right', fontFamily: 'B Nazanin', topBorder: 'false', bottomBorder: 'false', rightBorder: 'false', leftBorderSize: 17 } },
            { x: 33, y: 1, value: '', mergeRow: '', mergeCol: '' },
            { x: 33, y: 2, value: '', mergeRow: '', mergeCol: '' },
            { x: 33, y: 3, value: '', mergeRow: '', mergeCol: '' },
            { x: 33, y: 4, value: '', mergeRow: '', mergeCol: '' },
            { x: 33, y: 5, value: '11340', mergeRow: '', mergeCol: '', style: { align: 'right', fontFamily: 'B Nazanin', bottomBorder: 'false', leftBorder: 'false', topBorder: 'false' } },
            { x: 33, y: 6, value: '1124', mergeRow: '', mergeCol: '', style: { align: 'right', fontFamily: 'B Nazanin', bottomBorder: 'false', leftBorder: 'false', rightBorder: 'false', topBorder: 'false' } },
            { x: 33, y: 7, value: '7073', mergeRow: '', mergeCol: '', style: { align: 'right', fontFamily: 'B Nazanin', bottomBorder: 'false', leftBorder: 'false', rightBorder: 'false', topBorder: 'false' } },
            { x: 33, y: 8, value: '3143', mergeRow: '', mergeCol: '3', style: { align: 'right', fontFamily: 'B Nazanin', bottomBorder: 'false', leftBorder: 'false', rightBorder: 'false', topBorder: 'false' } },
            { x: 33, y: 9, value: '', mergeRow: '', mergeCol: '' },
            { x: 33, y: 10, value: '', mergeRow: '', mergeCol: '' },
            // //
            // //
            { x: 34, y: 0, value: 'گیلان -----------------------------------', mergeRow: '', mergeCol: '5', style: { align: 'right', fontFamily: 'B Nazanin', topBorder: 'false', bottomBorder: 'false', rightBorder: 'false', leftBorderSize: 17 } },
            { x: 34, y: 1, value: '', mergeRow: '', mergeCol: '' },
            { x: 34, y: 2, value: '', mergeRow: '', mergeCol: '' },
            { x: 34, y: 3, value: '', mergeRow: '', mergeCol: '' },
            { x: 34, y: 4, value: '', mergeRow: '', mergeCol: '' },
            { x: 34, y: 5, value: '15853', mergeRow: '', mergeCol: '', style: { align: 'right', fontFamily: 'B Nazanin', bottomBorder: 'false', leftBorder: 'false', topBorder: 'false' } },
            { x: 34, y: 6, value: '2092', mergeRow: '', mergeCol: '', style: { align: 'right', fontFamily: 'B Nazanin', bottomBorder: 'false', leftBorder: 'false', rightBorder: 'false', topBorder: 'false' } },
            { x: 34, y: 7, value: '9218', mergeRow: '', mergeCol: '', style: { align: 'right', fontFamily: 'B Nazanin', bottomBorder: 'false', leftBorder: 'false', rightBorder: 'false', topBorder: 'false' } },
            { x: 34, y: 8, value: '4543', mergeRow: '', mergeCol: '3', style: { align: 'right', fontFamily: 'B Nazanin', bottomBorder: 'false', leftBorder: 'false', rightBorder: 'false', topBorder: 'false' } },
            { x: 34, y: 9, value: '', mergeRow: '', mergeCol: '' },
            { x: 34, y: 10, value: '', mergeRow: '', mergeCol: '' },
            // //
            // //
            // //
            { x: 35, y: 0, value: 'لرستان -----------------------------------', mergeRow: '', mergeCol: '5', style: { align: 'right', fontFamily: 'B Nazanin', topBorder: 'false', bottomBorder: 'false', rightBorder: 'false', leftBorderSize: 17 } },
            { x: 35, y: 1, value: '', mergeRow: '', mergeCol: '' },
            { x: 35, y: 2, value: '', mergeRow: '', mergeCol: '' },
            { x: 35, y: 3, value: '', mergeRow: '', mergeCol: '' },
            { x: 35, y: 4, value: '', mergeRow: '', mergeCol: '' },
            { x: 35, y: 5, value: '10412', mergeRow: '', mergeCol: '', style: { align: 'right', fontFamily: 'B Nazanin', bottomBorder: 'false', leftBorder: 'false', topBorder: 'false' } },
            { x: 35, y: 6, value: '937', mergeRow: '', mergeCol: '', style: { align: 'right', fontFamily: 'B Nazanin', bottomBorder: 'false', leftBorder: 'false', rightBorder: 'false', topBorder: 'false' } },
            { x: 35, y: 7, value: '6786', mergeRow: '', mergeCol: '', style: { align: 'right', fontFamily: 'B Nazanin', bottomBorder: 'false', leftBorder: 'false', rightBorder: 'false', topBorder: 'false' } },
            { x: 35, y: 8, value: '2689', mergeRow: '', mergeCol: '3', style: { align: 'right', fontFamily: 'B Nazanin', bottomBorder: 'false', leftBorder: 'false', rightBorder: 'false', topBorder: 'false' } },
            { x: 35, y: 9, value: '', mergeRow: '', mergeCol: '' },
            { x: 35, y: 10, value: '', mergeRow: '', mergeCol: '' },
            // //
            // //
            // //
            { x: 36, y: 0, value: ' مازندران  ------------------------------', mergeRow: '', mergeCol: '5', style: { align: 'right', fontFamily: 'B Nazanin', topBorder: 'false', bottomBorder: 'false', rightBorder: 'false', leftBorderSize: 17 } },
            { x: 36, y: 1, value: '', mergeRow: '', mergeCol: '' },
            { x: 36, y: 2, value: '', mergeRow: '', mergeCol: '' },
            { x: 36, y: 3, value: '', mergeRow: '', mergeCol: '' },
            { x: 36, y: 4, value: '', mergeRow: '', mergeCol: '' },
            { x: 36, y: 5, value: '22463', mergeRow: '', mergeCol: '', style: { align: 'right', fontFamily: 'B Nazanin', bottomBorder: 'false', leftBorder: 'false', topBorder: 'false' } },
            { x: 36, y: 6, value: '2058', mergeRow: '', mergeCol: '', style: { align: 'right', fontFamily: 'B Nazanin', bottomBorder: 'false', leftBorder: 'false', rightBorder: 'false', topBorder: 'false' } },
            { x: 36, y: 7, value: '14638', mergeRow: '', mergeCol: '', style: { align: 'right', fontFamily: 'B Nazanin', bottomBorder: 'false', leftBorder: 'false', rightBorder: 'false', topBorder: 'false' } },
            { x: 36, y: 8, value: '5767', mergeRow: '', mergeCol: '3', style: { align: 'right', fontFamily: 'B Nazanin', bottomBorder: 'false', leftBorder: 'false', rightBorder: 'false', topBorder: 'false' } },
            { x: 36, y: 9, value: '', mergeRow: '', mergeCol: '' },
            { x: 36, y: 10, value: '', mergeRow: '', mergeCol: '' },
            // //
            // //
            { x: 37, y: 0, value: 'مرکزی  -----------------------------------', mergeRow: '', mergeCol: '5', style: { align: 'right', fontFamily: 'B Nazanin', topBorder: 'false', bottomBorder: 'false', rightBorder: 'false', leftBorderSize: 17 } },
            { x: 37, y: 1, value: '', mergeRow: '', mergeCol: '' },
            { x: 37, y: 2, value: '', mergeRow: '', mergeCol: '' },
            { x: 37, y: 3, value: '', mergeRow: '', mergeCol: '' },
            { x: 37, y: 4, value: '', mergeRow: '', mergeCol: '' },
            { x: 37, y: 5, value: '8304', mergeRow: '', mergeCol: '', style: { align: 'right', fontFamily: 'B Nazanin', bottomBorder: 'false', leftBorder: 'false', topBorder: 'false' } },
            { x: 37, y: 6, value: '946', mergeRow: '', mergeCol: '', style: { align: 'right', fontFamily: 'B Nazanin', bottomBorder: 'false', leftBorder: 'false', rightBorder: 'false', topBorder: 'false' } },
            { x: 37, y: 7, value: '4725', mergeRow: '', mergeCol: '', style: { align: 'right', fontFamily: 'B Nazanin', bottomBorder: 'false', leftBorder: 'false', rightBorder: 'false', topBorder: 'false' } },
            { x: 37, y: 8, value: '2633', mergeRow: '', mergeCol: '3', style: { align: 'right', fontFamily: 'B Nazanin', bottomBorder: 'false', leftBorder: 'false', rightBorder: 'false', topBorder: 'false' } },
            { x: 37, y: 9, value: '', mergeRow: '', mergeCol: '' },
            { x: 37, y: 10, value: '', mergeRow: '', mergeCol: '' },
            // //
            // //
            // //
            { x: 38, y: 0, value: 'همدان -----------------------------------', mergeRow: '', mergeCol: '5', style: { align: 'right', fontFamily: 'B Nazanin', topBorder: 'false', bottomBorder: 'false', rightBorder: 'false', leftBorderSize: 17 } },
            { x: 38, y: 1, value: '', mergeRow: '', mergeCol: '' },
            { x: 38, y: 2, value: '', mergeRow: '', mergeCol: '' },
            { x: 38, y: 3, value: '', mergeRow: '', mergeCol: '' },
            { x: 38, y: 4, value: '', mergeRow: '', mergeCol: '' },
            { x: 38, y: 5, value: '8972', mergeRow: '', mergeCol: '', style: { align: 'right', fontFamily: 'B Nazanin', bottomBorder: 'false', leftBorder: 'false', topBorder: 'false' } },
            { x: 38, y: 6, value: '937', mergeRow: '', mergeCol: '', style: { align: 'right', fontFamily: 'B Nazanin', bottomBorder: 'false', leftBorder: 'false', rightBorder: 'false', topBorder: 'false' } },
            { x: 38, y: 7, value: '5612', mergeRow: '', mergeCol: '', style: { align: 'right', fontFamily: 'B Nazanin', bottomBorder: 'false', leftBorder: 'false', rightBorder: 'false', topBorder: 'false' } },
            { x: 38, y: 8, value: '2423', mergeRow: '', mergeCol: '3', style: { align: 'right', fontFamily: 'B Nazanin', bottomBorder: 'false', leftBorder: 'false', rightBorder: 'false', topBorder: 'false' } },
            { x: 38, y: 9, value: '', mergeRow: '', mergeCol: '' },
            { x: 38, y: 10, value: '', mergeRow: '', mergeCol: '' },
            // //
            // //
            { x: 39, y: 0, value: 'هرمزگان ---------------------------------', mergeRow: '', mergeCol: '5', style: { align: 'right', fontFamily: 'B Nazanin', topBorder: 'false', bottomBorder: 'false', rightBorder: 'false', leftBorderSize: 17 } },
            { x: 39, y: 1, value: '', mergeRow: '', mergeCol: '' },
            { x: 39, y: 2, value: '', mergeRow: '', mergeCol: '' },
            { x: 39, y: 3, value: '', mergeRow: '', mergeCol: '' },
            { x: 39, y: 4, value: '', mergeRow: '', mergeCol: '' },
            { x: 39, y: 5, value: '12386', mergeRow: '', mergeCol: '', style: { align: 'right', fontFamily: 'B Nazanin', bottomBorder: 'false', leftBorder: 'false', rightBorder: 'false', topBorder: 'false' } },
            { x: 39, y: 6, value: '1150', mergeRow: '', mergeCol: '', style: { align: 'right', fontFamily: 'B Nazanin', bottomBorder: 'false', leftBorder: 'false', rightBorder: 'false', topBorder: 'false' } },
            { x: 39, y: 7, value: '8049', mergeRow: '', mergeCol: '', style: { align: 'right', fontFamily: 'B Nazanin', bottomBorder: 'false', leftBorder: 'false', rightBorder: 'false', topBorder: 'false' } },
            { x: 39, y: 8, value: '3187', mergeRow: '', mergeCol: '3', style: { align: 'right', fontFamily: 'B Nazanin', bottomBorder: 'false', leftBorder: 'false', rightBorder: 'false', topBorder: 'false' } },
            { x: 39, y: 9, value: '', mergeRow: '', mergeCol: '' },
            { x: 39, y: 10, value: '', mergeRow: '', mergeCol: '' },
            // //
            // //
            // //
            { x: 40, y: 0, value: ' یزد -----------------------------------', mergeRow: '', mergeCol: '5', style: { align: 'right', fontFamily: 'B Nazanin', topBorder: 'false', bottomBorder: 'false', rightBorder: 'false', leftBorderSize: 17 } },
            { x: 40, y: 1, value: '', mergeRow: '', mergeCol: '' },
            { x: 40, y: 2, value: '', mergeRow: '', mergeCol: '' },
            { x: 40, y: 3, value: '', mergeRow: '', mergeCol: '' },
            { x: 40, y: 4, value: '', mergeRow: '', mergeCol: '' },
            { x: 40, y: 5, value: '8422 ', mergeRow: '', mergeCol: '', style: { align: 'right', fontFamily: 'B Nazanin', bottomBorder: 'false', leftBorder: 'false', rightBorder: 'false', topBorder: 'false' } },
            { x: 40, y: 6, value: '890', mergeRow: '', mergeCol: '', style: { align: 'right', fontFamily: 'B Nazanin', bottomBorder: 'false', leftBorder: 'false', rightBorder: 'false', topBorder: 'false' } },
            { x: 40, y: 7, value: '4803', mergeRow: '', mergeCol: '', style: { align: 'right', fontFamily: 'B Nazanin', bottomBorder: 'false', leftBorder: 'false', rightBorder: 'false', topBorder: 'false' } },
            { x: 40, y: 8, value: '2729', mergeRow: '', mergeCol: '3', style: { align: 'right', fontFamily: 'B Nazanin', bottomBorder: 'false', leftBorder: 'false', rightBorder: 'false', topBorder: 'false' } },
            { x: 40, y: 9, value: '', mergeRow: '', mergeCol: '' },
            { x: 40, y: 10, value: '', mergeRow: '', mergeCol: '' },
        ]; // data
        var docx = new Docx_1.Docx('temp5.docx', 'outputUnitTestTmp/');
        docx.createTable(data);
        var out = docx.generate();
        if (out == false) {
            assert_1.fail("Don't create File of docx");
        }
        else {
            console.log("Create File of Docx");
            done();
        }
    }); ////  it فایل مربوط به كاركنان شاغل در دانشگاههاي علوم پزشكي بر حسب گروه شغلي
}); // describe
//# sourceMappingURL=template.js.map