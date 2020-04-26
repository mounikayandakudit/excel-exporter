import { Component, OnInit } from '@angular/core';
import { Lookup } from '../Lookup.model';
import { ExcelService } from '../excel.service';

type lookup = {};

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


  constructor(private excelservice: ExcelService) { }

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
    let lookup;
    for (let prop in c) {
      lookup[prop] = c[prop];
    }
    return lookup;
  }
}


