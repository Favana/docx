import {} from 'mocha';
let json2xml = require('json2xml');
let _ = require('lodash');

export class table{

    createTr(_body:any, data:any){

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


    tableStyle(_body:any, style?:any, data?:any){
        let defaultStyle = {
            fontFamily : 'B Nazanin',
            fontSize : 20,
            fontColor : 'black',
            bold: 'false',
            direction :'rtl',
            align : 'right'
        };
        _body;

        if(style != null){
            let keysDefualtStyle = Object.keys(defaultStyle);
            for(let i=0; i<keysDefualtStyle.length; i++){
                if(style[keysDefualtStyle[i]] != undefined){
                    defaultStyle[keysDefualtStyle[i]] = style[keysDefualtStyle[i]];
                    defaultStyle;
                }
            }
        }else{
            defaultStyle;
        }

       _body;




        // {'w:rPr':[
        //     {'w:rtl':[]},
        //     {'w:rFonts':'', attr:{'w:cs':defaultStyle.fontFamily,'w:hint':'cs'}},
        //     {'w:szCs':'', attr:{'w:val':2*(defaultStyle.fontSize)}},
        //     {'w:color':'', attr:{'w:val':defaultStyle.fontColor}},
        // ]},


        // _body['w:tbl'][0]['w:tblPr'].push({'w:tblBorders':[ //<w:top w:val="single" w:sz="12" w:space="0" w:color="95B3D7" w:themeColor="accent1" w:themeTint="99"/>
        //         {'w:top':'', attr:{'w:val':"single",  'w:sz':"12",'w:color':'95B3D7', 'w:themeColor':'accent1', 'w:themeTint':'99'}}
        //     ]});

    }


    callingMethod(body:any, data:any, style?:any){
        let objTr = this.createTr(body, data);
        let objTc = this.createTc(objTr, data);
        let objP = this.createPtable(objTc, data);
        let objMerge = this.createMerge(objP, data);
        let objStyle = this.tableStyle(objMerge, style, data);
        return objMerge;
       // next(json2xml(objMerge,{ attributes_key:'attr' } ));
    }


}//  class


