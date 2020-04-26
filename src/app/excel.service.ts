import { Injectable } from '@angular/core';
import { HttpClient, HttpHeaders } from '@angular/common/http';

@Injectable({
  providedIn: 'root'
})
export class ExcelService {

  constructor(private http: HttpClient) { }

  public readStaticXlsx() {
    return this.http.get('../assets/TIBLOOKUPDATA.xlsx', {
      responseType: 'blob' // <-- changed to blob
    });
  }
}
