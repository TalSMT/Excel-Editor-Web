import { Component, OnInit } from '@angular/core';

@Component({
  selector: 'app-home',
  templateUrl: './home.component.html',
  styleUrls: ['./home.component.scss']
})
export class HomeComponent{
  title = 'Excel-File-App';
  fileName:any;
  constructor() { }

  DataFromEventEmitter(data:any) {
  
    console.log(data);
    // this.fileName=data.target.file;
  }

}
