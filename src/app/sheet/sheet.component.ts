import { Component, OnInit } from '@angular/core';
import { Lookup } from '../Lookup.model';
import { ExcelService } from '../excel.service';
import * as FileSaver from 'file-saver';
import * as XLSX from 'xlsx';


type lookup = {};
const EXCEL_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=UTF-8';
const EXCEL_EXTENSION = '.xlsx';

@Component({
  selector: 'app-sheet',
  templateUrl: './sheet.component.html',
  styleUrls: ['./sheet.component.css']
})
export class SheetComponent implements OnInit {


  displayDialog: boolean;

  lookup: Lookup;

  selectedLookup: Lookup;

  newLookup: boolean;

  lookups: Lookup[];

  cols: any[];


  constructor() { 
    
  }

  ngOnInit() {
    //this. excelservice.exportAsExcelFile().then(lookups=> this.lookups = this.lookups)

    this.cols = [
      { field: 'LOOKUP_ID', header: 'LOOKUP_ID' },
      { field: 'LOOKUP_CATEGORY', header: 'LOOKUP_CATEGORY' },
      { field: 'LOOKUP_SUBCATEGORY1', header: 'LOOKUP_SUBCATEGORY1' },
      { field: 'LOOKUP_SUBCATEGORY2', header: 'LOOKUP_SUBCATEGORY2' },
      { field: 'FROM_VALUE1', header: 'FROM_VALUE1' },
      { field: 'FROM_VALUE2', header: 'FROM_VALUE2' },
      { field: 'FROM_VALUE3', header: 'FROM_VALUE3' },
      { field: 'FROM_VALUE4', header: 'FROM_VALUE4' },
      { field: 'FROM_VALUE5', header: 'FROM_VALUE5' },
      { field: 'TO_VALUE1', header: 'TO_VALUE1' },
      { field: 'TO_VALUE2', header: 'TO_VALUE2' },
      { field: 'TO_VALUE3', header: 'TO_VALUE3' },
      { field: 'TO_VALUE4', header: 'TO_VALUE4' },
      { field: 'TO_VALUE5', header: 'TO_VALUE5' },
      { field: 'COMMENTS', header: 'COMMENTS' }
    ];

    this.lookups = [
      {'LOOKUP_ID': 1,'LOOKUP_CATEGORY':'aa','LOOKUP_SUBCATEGORY1':'ww','LOOKUP_SUBCATEGORY2':'ww','FROM_VALUE1':'ff','FROM_VALUE2':'ff','FROM_VALUE3':'ff','FROM_VALUE4':'ff',
      'FROM_VALUE5':'ff','TO_VALUE1':'aa','TO_VALUE2':'aa','TO_VALUE3':'aa','TO_VALUE4':'aa','TO_VALUE5':'aa','COMMENTS':'aa'},
      {'LOOKUP_ID': 2,'LOOKUP_CATEGORY':'aa','LOOKUP_SUBCATEGORY1':'ww','LOOKUP_SUBCATEGORY2':'ww','FROM_VALUE1':'ff','FROM_VALUE2':'ff','FROM_VALUE3':'ff','FROM_VALUE4':'ff',
      'FROM_VALUE5':'ff','TO_VALUE1':'aa','TO_VALUE2':'aa','TO_VALUE3':'aa','TO_VALUE4':'aa','TO_VALUE5':'aa','COMMENTS':'aa'}
    ];

  }
  showDialogToAdd() {
    this.newLookup = true;
    this.lookup = new Lookup();
    this.displayDialog = true;
    //this. excelservice.exportAsExcelFile().then(lookups=> this.lookups = this.lookups)

  }

  save() {
    let lookups = [...this.lookups];
    if (this.newLookup)
      lookups.push(this.lookup);
    else
      lookups[this.lookups.indexOf(this.selectedLookup)] = this.lookup;

    this.lookups = this.lookups;
    this.lookup = null;
    this.displayDialog = false;
    //this. excelservice.exportAsExcelFile(this.data,"sheet");
  }

  delete() {
    let index = this.lookups.indexOf(this.selectedLookup);
    this.lookups = this.lookups.filter((val, i) => i != index);
    this.lookup = null;
    this.displayDialog = false;
    //this. excelservice.exportAsExcelFile(this.data(),"sheet");

  }
  data(data: any) {
    throw new Error("Method not implemented.");
  }

  onRowSelect(event) {
    this.newLookup = false;
    this.lookup = this.cloneLookup(event.data);
    this.displayDialog = true;
  }

  cloneLookup(c: Lookup): Lookup {
    let lookup = new Lookup();
    for (let prop in c) {
      lookup[prop] = c[prop];
    }
    return lookup;
  }
  public exportAsExcelFile(json: any[], excelFileName: string): void {
    const worksheet: XLSX.WorkSheet = XLSX.utils.json_to_sheet(json);
    const workbook: XLSX.WorkBook = { Sheets: { 'data': worksheet }, SheetNames: ['data'] };
    const excelBuffer: any = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
    this.saveAsExcelFile(excelBuffer, excelFileName);
  }
  private saveAsExcelFile(buffer: any, fileName: string): void {
     const data: Blob = new Blob([buffer], {type: EXCEL_TYPE});
     FileSaver.saveAs(data, fileName + '_export_' + new  Date().getTime() + EXCEL_EXTENSION);
  }
  onFileChange(evt: any) {
    /* wire up file reader */
    const target: DataTransfer = <DataTransfer>(evt.target);
    if (target.files.length !== 1) throw new Error('Cannot use multiple files');
    const reader: FileReader = new FileReader();
    reader.onload = (e: any) => {
      /* read workbook */this.newLookup
      const bstr: string = e.target.result;
      const wb: XLSX.WorkBook = XLSX.read(bstr, {type: 'binary'});

      /* grab first sheet */
      const wsname: string = wb.SheetNames[0];
      const ws: XLSX.WorkSheet = wb.Sheets[wsname];

      this.lookups = [];
      /* save data */
      let data = (XLSX.utils.sheet_to_json(ws, {header: 1}));
      for(let row of data){
          console.log(row);
          let lookUp = new Lookup();
          lookUp.LOOKUP_ID = row[0];
          lookUp.LOOKUP_CATEGORY = row[1];
          this.lookups.push(lookUp);
      }
    };
    reader.readAsBinaryString(target.files[0]);
  }
}


