import { Directive, HostListener ,Output, EventEmitter,Input } from '@angular/core';
import { Observable, Subscriber } from 'rxjs';
import * as XLSX from 'xlsx';

@Directive({
  selector: '[appReadexcel]',
  exportAs:'readexcel',
  
})

export class ReadexcelDirective {
   excelObsrvable!:Observable<any>;
  @Output () eventEmitter = new EventEmitter();
 
  


  constructor() { }

  @HostListener("change",["$event.target"])
  onChange(target:HTMLInputElement){
     const file= target.files![0]; //בחרנו קובץ ושמרנו ב file
     
    
    this.excelObsrvable= new  Observable((subscriber:Subscriber<any>)=>{
    this.readFile(file,subscriber); //נקרא לפונקציה שתקרא את הקובץ

   });
 

   this.excelObsrvable.subscribe((d)=> {
     this.eventEmitter.emit(d);
   });

   
  }

  readFile(file:File,subscriber:Subscriber<any>){
    const fileReader = new FileReader();
    fileReader.readAsArrayBuffer(file);

    fileReader.onload=(e)=>{                 //נקבל את הדטה מהקובץ

    const bufferArray =e.target!.result;      //נקבל פה את המערך שבבאפר 
    const wb: XLSX.WorkBook= XLSX.read(bufferArray, {type:'buffer'});
    const wsname:string = wb.SheetNames[0];      //השם של הגיליון הראשון שאותו רוצים
    const ws :XLSX.WorkSheet= wb.Sheets[wsname]      //לקרוא את הגיליון הזה
    const data =XLSX.utils.sheet_to_json(ws);

    subscriber.next(data);
    subscriber.complete();




    }

  }
  


}
