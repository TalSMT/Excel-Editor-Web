import { KeyValue } from '@angular/common';
import { Component, Input, ViewChild ,OnChanges, SimpleChange, SimpleChanges} from '@angular/core';
import * as XLSX from 'xlsx';
import { ReadexcelDirective } from '../directives/readexcel.directive';

@Component({
  selector: 'app-writeexcel',
  templateUrl: './writeexcel.component.html',
  styleUrls: ['./writeexcel.component.scss'],
  
})

export class WriteexcelComponent implements OnChanges{

  title = 'angular-app';
  fileName= 'ExcelSheet.xlsx';
  dataFile:any;
  dataSheet :any;
  isLoading = false;

  

  constructor() {  }
  //  func1 (record:any){
  //   var net =(record.Num1) - (record.Num2);
  //   record.Net = net;
  //   var n :string= (<HTMLInputElement>document.getElementById("nameOfCol")).value;
  //   delete record[n];
  //   delete record.Num2;
  //   return record;

  // }


 addCol(record:any){
   console.log("hi");
   var nameNewCol:string = (<HTMLInputElement>document.getElementById("addNewCol")).value;
   var textCol :string =  (<HTMLInputElement>document.getElementById("text")).value;
   console.log(nameNewCol+"nameNewCol");
   record[nameNewCol]=textCol;
   return record;
 }

///------double a column by fixed number
  doubleCol (record:any){
    var fix =(<HTMLInputElement>document.getElementById("fixNum")).value;
    if(fix == "") return;
    var fixNum:number =+fix;
    var col = (<HTMLInputElement>document.getElementById("theColToDouble")).value;
    var newCol = (<HTMLInputElement>document.getElementById("nameOfNewCol")).value;
    
        var el= fixNum  * record[col];
        // record[col]=el;
     
        record[newCol] = el; 
           
    return record;

  }



///------delete column
  delete_Col(record:any){ 
    var n :string= (<HTMLInputElement>document.getElementById("nameOfCol")).value;
    if(n == null) return;
    var separators = /\s*(?:,|'')\s*/ 
    n.split(separators).forEach((item)=>{
      console.log(item+"item");
      console.log(record[item]+"recitem");
      delete record[item];
      
    })
   
    
    // delete record[n];    
    
    return record;
  }


///-----delete one row
ec (r:number, c:number){
  return XLSX.utils.encode_cell({r:r,c:c});
 }

delete_row(ws:any, row_index:number){
//  var row_index = (<HTMLInputElement>document.getElementById("numOfRow")).value;
 var variable = XLSX.utils.decode_range(ws['!ref']);
 var R = +row_index;
  console.log(R+"R");
  for(R; R < variable.e.r; ++R){
      for(var C = variable.s.c; C <= variable.e.c; ++C){
          ws[this.ec(R,C)] = ws[this.ec(R+1,C)];
      }
  }

  variable.e.r--
  
  ws['!ref'] = XLSX.utils.encode_range(variable.s, variable.e);
   return ws;
}

// delete range of rows

delete_range_rows(ws:any, rows_index:string){

  // var rows_index = (<HTMLInputElement>document.getElementById("numOfRowE")).value;
 
  console.log("1");
  var variable = XLSX.utils.decode_range(ws['!ref']);
  console.log("2");
  var rangeArr = new Array();
      rangeArr = rows_index.split("-");
	  console.log(rangeArr+"rangeArr");
    console.log(rangeArr[0]+"rangeArr[0]");
    console.log(rangeArr[1]+"rangeArr[1]");
    

  //  var separators2 = /\s*(?:-)\s*/ 
  var delta:number = +Math.abs(rangeArr[0] - rangeArr[1])+1; 
  console.log(delta+"delta");
  var a1 :number = rangeArr[1] > rangeArr[0] ? +rangeArr[0] : +rangeArr[1] ;
  console.log(a1+"a1");
  var c:number=delta+a1;
  console.log(c+"a1+d");
  
  
 
   
   for(var R=a1; R < variable.e.r; ++R){
      for(var C = variable.s.c; C <= variable.e.c; ++C){
           ws[this.ec(R,C)] = ws[this.ec(R+delta,C)];
      }
    }
	 variable.e.r--
    
    ws['!ref'] = XLSX.utils.encode_range(variable.s, variable.e);
     return ws;
     
}






///ניסיון לעבור על מה שפיצלנו ולמחוק את הערכים שמופרדים בפסיקים
// var separators1 = /\s*(?:,|'')\s*/ 
// row_index.split(separators1).forEach((item)=>{
//   console.log(item+"i");
// var variable = XLSX.utils.decode_range(ws['!ref']);
//  var R = +item;
//  console.log(R+"R");
//  console.log(variable.e.r+"variable.e.r");
//   for(R; R < variable.e.r; ++R){
//       for(var C = variable.s.c; C <= variable.e.c; ++C){
//           ws[this.ec(R,C)] = ws[this.ec(R+1,C)];
//       }
//   }
//   variable.e.r--
  
//   ws['!ref'] = XLSX.utils.encode_range(variable.s, variable.e);
 
//  })



// for(R; R < variable.e.r; ++R){
//   for(var C = variable.s.c; C <= variable.e.c; ++C){
//       ws[this.ec(R,C)] = ws[this.ec(R+delta,C)];
//   }
// }

// delete_row(record:any){
//   var n :string= (<HTMLInputElement>document.getElementById("numOfRow")).value;
//   console.log(n+"n");
//   // var t : number = +n;
//   //  console.log(record[]);
  
// //  delete record[t-1];
//   return record;
// }


delete_colByIndex(record:any){
  // var n :string= (<HTMLInputElement>document.getElementById("numOfRow")).value;
  // var t : number = +n;
  // delete record[Object.keys(record)[t-1]];
  return record;
}

////-----main function that call all the functions that change the file-------------
changeFile(){


/// -----------call function that delete a column
 var fixs =(<HTMLInputElement>document.getElementById("fixNum")).value;
 var n = (<HTMLInputElement>document.getElementById("nameOfCol")).value;
 var nameNewCol = (<HTMLInputElement>document.getElementById("addNewCol")).value;
 var rowIndex = (<HTMLInputElement>document.getElementById("numOfRow")).value;
 var rowRange = (<HTMLInputElement>document.getElementById("numOfRange")).value;


  if (n != ""){
      this.dataFile = this.dataFile.map(this.delete_Col);
    }
  if (fixs != ""){
  this.dataFile = this.dataFile.map(this.doubleCol);
    }

  if (nameNewCol!= ""){
  this.dataFile = this.dataFile.map(this.addCol);
    }

  this.dataSheet = XLSX.utils.json_to_sheet(this.dataFile);

  ///----------call function that delete a row

  // create a var thet store the input value of the row that want to delete
  
  if (rowIndex != ""){
  var r_convert: number = +rowIndex;
  this.dataSheet= this.delete_row(this.dataSheet,r_convert);
  }

  if(rowRange != ""){
    // var re_convert : number = +rowRange;
    this.dataSheet = this.delete_range_rows(this.dataSheet, rowRange);
  }

  return this.dataSheet;

}




ngOnChanges(change: SimpleChanges){
  console.log(change + "kkk");
}


///-----------------function that write the file to excel--------------
  exportexcel(): void {

    this.isLoading = true;
    setTimeout(()=>{
      this.isLoading=false;
    },2000)
    
    /* pass here the table id */
    let element = document.getElementById('thetable');
  
    const ws: XLSX.WorkSheet =XLSX.utils.table_to_sheet(element);
     this.dataFile = XLSX.utils.sheet_to_json(ws);
    
  
    //  this.dataFile= this.dataFile.map(this.delete_Col);
  
    
    
    // var newData =  this.dataFile.map((record:any)=>{
    //   // var net =(record.Num1) - (record.Num2);
    //   // record.Net = net;
    //   // delete record.Num1;
    //   var n ='Num2';
    //   delete record[n];
    //   return record;

    // });
   
 
    /* generate workbook and add the worksheet */
    const wb: XLSX.WorkBook = XLSX.utils.book_new();
    // var newWs = XLSX.utils.json_to_sheet(this.dataFile);
    XLSX.utils.book_append_sheet(wb, this.changeFile(), 'Sheet1');
 
    /* save to file */  
    XLSX.writeFile(wb, this.fileName);
 
  }

}
