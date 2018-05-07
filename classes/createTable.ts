import {} from 'mocha';
import {type} from "os";
let json2xml = require('json2xml');
let _ = require('lodash');

export class table{

    createTr(_body:any, data:any){
        _body;
        let check = _body['w:tbl'].length;
        if(check == 1){
             _body ;
        }else{
            for(let i=1; i<check; i++){
                _body['w:tbl'].pop();
            }
            //this.globalTbl = (<any>Object).assign(resultTable, this.globalTbl);
            // let p = {'w:p':[]};
            // _body = (<any>Object).assign(p, _body);
            _body;
          //  return _body;
        }
        _body;
        let counterRow =_.uniqBy(data, 'x');
        let row = counterRow.length;

        for(let i=0; i<row; i++ ){
            _body['w:tbl'].push({'w:tr':[]});
        }
        return _body;

    }//createTr


    createTc(_body:any, data:any){

        let counterCol = _.uniqBy(data, 'y');
        let counterRow =_.uniqBy(data, 'x');
        let row = counterRow.length;
        let col = counterCol.length;

        for(let i=1; i<=row; i++){
            for(let j=0; j<col; j++){
                _body['w:tbl'][i]['w:tr'].push({'w:tc':[{'w:tcPr':[{'w:tcW':'' , attr:{'w:w':4657, 'w:type':'dxa'}}]},]});
            }// for column
        }//for row
        return _body;
    }//createTc

    createPtable(_body:any, data:any){

        let counterCol = _.uniqBy(data, 'y');
        let counterRow =_.uniqBy(data, 'x');
        let row = counterRow.length;
        let col = counterCol.length;

        for(let i=1; i<=row; i++){
            for(let j=0; j<col; j++){
                let checkValue = _.find(data, {x:i,y:j});
                let value = checkValue.value;
                if(value == ''){
                    _body['w:tbl'][i]['w:tr'][j]['w:tc'].push({'w:p':[]});
                }else{
                    _body['w:tbl'][i]['w:tr'][j]['w:tc'].push({'w:p':[

                        {'w:r':[
                                {'w:t':value}
                            ]}]});
                }
            }// for col
        }// for row

        return _body;
    }//createP



    createMerge(_body:any,data:any){

        let counterCol = _.uniqBy(data, 'y');
        let counterRow =_.uniqBy(data, 'x');
        let row = counterRow.length;
        let col = counterCol.length;
        let start =[];
        for(let i=1;i<=row; i++ ){
            for(let j=0; j<col;j++ ){
                let checkMerge = _.find(data, {x:i, y:j});
                let rowX= checkMerge.mergeRow;
                let colY = checkMerge.mergeCol;

                if( colY != '' && rowX !='' ){
                    let myX = checkMerge.x;
                    let myY = checkMerge.y;
                    let counterY = colY;// mergeCol
                    let counterX = rowX; //mergeRow
                    /// start merge Row and Col
                    _body['w:tbl'][myX]['w:tr'][myY]['w:tc'][0]['w:tcPr'].push({'w:vMerge':'', attr:{'w:val':'restart'}},{'w:gridSpan':'', attr:{'w:val': colY}}); ///<w:vMerge w:val="restart"/>
                    let b= _body;
                    ///// loop for merge Col:
                    for(let i=1; i<counterY; i++){
                        myY= myY+1;
                        delete _body['w:tbl'][myX]['w:tr'][myY]; // merge col
                    }// for
                    /// end loop for
                    let c = _body;
                    ///// loop for merge Row:
                    for(let i =1; i<counterX; i++){
                        myY =0;
                        myY = checkMerge.y
                        myX=myX+1;
                        _body['w:tbl'][myX]['w:tr'][myY]['w:tc'][0]['w:tcPr'].push({'w:vMerge':[]},{'w:gridSpan':'', attr:{'w:val': colY}}); // merge row
                    }// for
                    /// end loop for
                    let z =_body;
                    ///// loop for merge Col:
                    for(let i=1; i<counterY; i++){
                        myY = myY+1;
                        delete _body['w:tbl'][myX]['w:tr'][myY]; // merge col
                    }// for
                    /// end loop for
                    let a =_body;

                }// if


                else if(colY !='' && rowX ==''){
                    let myX = checkMerge.x;
                    let myY = checkMerge.y;
                    let counterY = colY;

                    /// start merge Row and Col
                    _body['w:tbl'][myX]['w:tr'][myY]['w:tc'][0]['w:tcPr'].push({'w:gridSpan':'', attr:{'w:val': colY}}); ///<w:gridSpan w:val="2"/>
                    let b= _body;
                    ///// loop for merge Col:
                    for(let i=1 ; i<counterY; i++){
                        myY = myY+1;
                        delete _body['w:tbl'][myX]['w:tr'][myY]; // merge col
                    }// for
                    let c = _body;

                }// else if


                else if(colY =='' && rowX != ''){
                    let myY = checkMerge.y;
                    let myX = checkMerge.x;
                    let counterX = rowX;

                    /// start merge Row and Row
                    _body['w:tbl'][myX]['w:tr'][myY]['w:tc'][0]['w:tcPr'].push({'w:vMerge':'', attr:{'w:val':'restart'}}); ///<w:vMerge w:val="restart"/>
                    let b= _body;
                    ///// loop for merge Row:
                    for(let i=1; i<counterX; i++){
                        myX = myX+1;
                        _body['w:tbl'][myX]['w:tr'][myY]['w:tc'][0]['w:tcPr'].push({'w:vMerge':[]}); // merge row
                    }// for
                    let c = _body;

                } // else ifs



            }// for col
        }// for row
        return _body

    }//createMerge


    tableStyle(_body:any, style?:any){
       style;
       _body;

        if(style != null ){
            if(typeof  style == 'object'){
                let counterX = _.uniqBy(style, 'x');
                let counterY =  _.uniqBy(style, 'y');
                let counterCol = counterY.length;
                let counterRow = counterX.length;

                for(let i= 1; i<=counterRow; i++){
                    for(let j= 0; j<counterCol; j++){
                        let find = _.find(style, {x:i, y:j});
                        let borderSize = find.borderSize;
                        let align = find.align;
                        let fontFamily = find.fontFamily;
                        let fontColor = find.fontColor;
                        let fontSize = find.fontSize;
                        let backgroundCell = find.background;
                        let bold = find.bold;


                        if(borderSize != undefined){ ///// if check sizeBorder
                            _body['w:tbl'][i]['w:tr'][j]['w:tc'][0]['w:tcPr'].push(
                                {'w:tblBorders':[
                                        {'w:top':'', attr:{'w:val':'single', 'w:sz':borderSize, 'w:space':'0', 'w:color':'auto'}},
                                        {'w:left':'', attr:{'w:val':'single', 'w:sz':borderSize, 'w:space':'0', 'w:color':'auto'}},
                                        {'w:bottom':'', attr:{'w:val':'single', 'w:sz':borderSize, 'w:space':'0', 'w:color':'auto'}},
                                        {'w:right':'', attr:{'w:val':'single', 'w:sz':borderSize, 'w:space':'0', 'w:color':'auto'}}
                                 ]}// 'w:tblBorders'

                            );
                            _body;
                        }// if check sizeBorder

                        if(align != undefined){ // if check align
                            let check = _body['w:tbl'][i]['w:tr'][j]['w:tc'][1]['w:p'].length;

                            if(check == 1){
                                _body['w:tbl'][i]['w:tr'][j]['w:tc'][1]['w:p'].splice(0, 0 ,
                                    {'w:pPr':[
                                            {'w:jc':'',attr:{'w:val':align}}
                                        ]}
                                );
                                _body;
                            }else if (check > 1){
                                _body['w:tbl'][i]['w:tr'][j]['w:tc'][1]['w:p'][0]['w:pPr'].push(
                                            {'w:jc':'',attr:{'w:val':align}}
                                );
                                _body;
                            }
                        }// if check align




                        if(fontFamily != undefined){ // if  check font
                            let checkP = _body['w:tbl'][i]['w:tr'][j]['w:tc'][1]['w:p'].length;

                            if(checkP == 1){
                                let checkR = _body['w:tbl'][i]['w:tr'][j]['w:tc'][1]['w:p'][0]['w:r'].length;

                                if(checkR == 1){
                                    _body['w:tbl'][i]['w:tr'][j]['w:tc'][1]['w:p'][0]['w:r'].splice(0, 0 ,
                                        {'w:rPr':[
                                                {'w:rFonts':'' , attr:{'w:cs':fontFamily}},
                                                {'w:rtl':''}
                                            ]}// 'w:rPr'
                                    );
                                    _body;
                                }else if(checkR > 1){
                                    _body['w:tbl'][i]['w:tr'][j]['w:tc'][1]['w:p'][0]['w:r'][0]['w:rPr'].push(
                                        {'w:rFonts':'' , attr:{'w:cs':fontFamily}},
                                        {'w:rtl':''}
                                    );
                                    _body;
                                }
                            }else if(checkP>1){
                                let checkR = _body['w:tbl'][i]['w:tr'][j]['w:tc'][1]['w:p'][1]['w:r'].length;

                                if(checkR == 1){
                                    _body['w:tbl'][i]['w:tr'][j]['w:tc'][1]['w:p'][1]['w:r'].splice(0, 0 ,
                                        {'w:rPr':[
                                                {'w:rFonts':'' , attr:{'w:cs':fontFamily}},
                                                {'w:rtl':''}
                                            ]}// 'w:rPr'
                                    );
                                    _body;
                                }else if(checkR > 1){
                                    _body['w:tbl'][i]['w:tr'][j]['w:tc'][1]['w:p'][1]['w:r'][0]['w:rPr'].push(
                                        {'w:rFonts':'' , attr:{'w:cs':fontFamily}},
                                        {'w:rtl':''}
                                    );
                                    _body;
                                }
                            }
                        }// if  check font



                        if(fontColor != undefined){ // if  fontColor
                            let checkP = _body['w:tbl'][i]['w:tr'][j]['w:tc'][1]['w:p'].length;
                            if(checkP == 1){
                                let checkR = _body['w:tbl'][i]['w:tr'][j]['w:tc'][1]['w:p'][0]['w:r'].length;
                                if(checkR == 1){
                                    _body['w:tbl'][i]['w:tr'][j]['w:tc'][1]['w:p'][0]['w:r'].splice(0,0,
                                        {'w:rPr':[
                                                {'w:color':'' , attr:{'w:val':fontColor}}
                                            ]}// 'w:rPr'
                                    );
                                    _body;
                                }else if(checkR >1){
                                    _body['w:tbl'][i]['w:tr'][j]['w:tc'][1]['w:p'][0]['w:r'][0]['w:rPr'].push(
                                        {'w:color':'' , attr:{'w:val':fontColor}}
                                    );
                                    _body;
                                }

                            }else if(checkP >1){
                                let checkR = _body['w:tbl'][i]['w:tr'][j]['w:tc'][1]['w:p'][1]['w:r'].length;
                                if(checkR == 1){
                                    _body['w:tbl'][i]['w:tr'][j]['w:tc'][1]['w:p'][1]['w:r'].splice(0,0,
                                        {'w:rPr':[
                                                {'w:color':'' , attr:{'w:val':fontColor}}
                                            ]}// 'w:rPr'
                                    );
                                    _body;
                                }else if(checkR >1){
                                    _body['w:tbl'][i]['w:tr'][j]['w:tc'][1]['w:p'][1]['w:r'][0]['w:rPr'].push(
                                        {'w:color':'' , attr:{'w:val':fontColor}}
                                    );
                                    _body;
                                }
                            }
                        }// if check fontColor


                        if(fontSize != undefined){ // if check fontSize
                            let checkP = _body['w:tbl'][i]['w:tr'][j]['w:tc'][1]['w:p'].length;
                            fontSize = 2*fontSize;
                            if(checkP == 1){
                                let checkR = _body['w:tbl'][i]['w:tr'][j]['w:tc'][1]['w:p'][0]['w:r'].length;
                                if(checkR == 1){
                                    _body['w:tbl'][i]['w:tr'][j]['w:tc'][1]['w:p'][0]['w:r'].splice(0, 0 ,
                                        {'w:rPr':[
                                                {'w:sz':'' , attr:{'w:val':fontSize}}
                                            ]}// 'w:rPr'
                                    );
                                    _body;
                                }else if(checkR > 1 ){
                                    _body['w:tbl'][i]['w:tr'][j]['w:tc'][1]['w:p'][0]['w:r'][0]['w:rPr'].push(
                                        {'w:sz':'' , attr:{'w:val':fontSize}}
                                    );
                                    _body;
                                }
                            }else if(checkP >1){
                                let checkR = _body['w:tbl'][i]['w:tr'][j]['w:tc'][1]['w:p'][1]['w:r'].length;
                                if(checkR == 1){
                                    _body['w:tbl'][i]['w:tr'][j]['w:tc'][1]['w:p'][1]['w:r'].splice(0, 0 ,
                                        {'w:rPr':[
                                                {'w:sz':'' , attr:{'w:val':fontSize}}
                                            ]}// 'w:rPr'
                                    );
                                    _body;
                                }else if(checkR > 1 ){
                                    _body['w:tbl'][i]['w:tr'][j]['w:tc'][1]['w:p'][1]['w:r'][0]['w:rPr'].push(
                                        {'w:sz':'' , attr:{'w:val':fontSize}}
                                    );
                                    _body;
                                }


                            }
                        }//  if check fontSize

                       if(backgroundCell != undefined){ // if check background cell table
                            _body['w:tbl'][i]['w:tr'][j]['w:tc'][0]['w:tcPr'].push(
                                {'w:shd':'' , attr:{'w:val':'clear','w:color':'auto', 'w:fill': backgroundCell}}
                            );
                           _body;
                       }// if check background cell table


                        if( bold != undefined ){ // if for check bold
                            let checkP = _body['w:tbl'][i]['w:tr'][j]['w:tc'][1]['w:p'].length;
                            if(checkP == 1){
                                let checkR = _body['w:tbl'][i]['w:tr'][j]['w:tc'][1]['w:p'][0]['w:r'].length;
                                if(checkR == 1 ){
                                    _body['w:tbl'][i]['w:tr'][j]['w:tc'][1]['w:p'][0]['w:r'].splice(0,0,
                                        {'w:rPr':[
                                                {'w:bCs':''}  // <w:bCs/>
                                            ]}// 'w:rPr'
                                    );
                                    _body;
                                }else if(checkR >1){
                                    _body['w:tbl'][i]['w:tr'][j]['w:tc'][1]['w:p'][0]['w:r'].push(
                                        {'w:bCs':''}  // <w:bCs/>
                                    );
                                    _body;
                                }

                            }else if(checkP > 1){
                                let checkR = _body['w:tbl'][i]['w:tr'][j]['w:tc'][1]['w:p'][1]['w:r'].length;
                                if(checkR == 1 ){
                                    _body['w:tbl'][i]['w:tr'][j]['w:tc'][1]['w:p'][1]['w:r'].splice(0,0,
                                        {'w:rPr':[
                                                {'w:bCs':''}  // <w:bCs/>
                                            ]}// 'w:rPr'
                                    );
                                    _body;
                                }else if(checkR >1){
                                    _body['w:tbl'][i]['w:tr'][j]['w:tc'][1]['w:p'][1]['w:r'].push(
                                        {'w:bCs':''}  // <w:bCs/>
                                    );
                                    _body;
                                }

                            }
                        }// if for check bold

                    }// for j
                }// for i
                _body;

            }else{
                throw "Type of style is not object";
            }// else

        }else{
            return _body;
        }//  else


    }


    callingMethod(body:any, data:any, style?:any){
        let objTr = this.createTr(body, data);
        let objTc = this.createTc(objTr, data);
        let objP = this.createPtable(objTc, data);
        let objMerge = this.createMerge(objP, data);
        let objStyle = this.tableStyle(objMerge, style);
        return objMerge;
       // next(json2xml(objMerge,{ attributes_key:'attr' } ));
    }


}//  class


