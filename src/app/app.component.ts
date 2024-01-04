import { Component, ElementRef, ViewChild } from '@angular/core';
import * as XLSX from 'xlsx';
import { saveAs } from 'file-saver';

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.scss'],
})
export class AppComponent {
  @ViewChild('addBtn', { static: false }) addBtn: ElementRef | undefined;
  @ViewChild('fileInput', { static: false }) fileInput: ElementRef | undefined;
  public updatedExcelObj: any = null;

  public uploadExcel() {
    this.fileInput?.nativeElement.click();
  }

  public getExcel(event: any) {
    let file = event.target.files[0];
    const fileName = file.name.slice(0, file.name.lastIndexOf('.'));
    const array = file.name.split('.');
    const extension = array[array.length - 1].toLowerCase();
    if (extension !== 'xls' && extension !== 'xlsx') {
      alert('Please Upload Excel File');
      return;
    }
    let fileReader = new FileReader();
    let summedData: any = {};
    fileReader.readAsBinaryString(file);
    fileReader.onload = (e) => {
      let workBook = XLSX.read(fileReader.result, { type: 'binary' });
      let sheetNames = workBook.SheetNames;
      let data = XLSX.utils.sheet_to_json(workBook.Sheets[sheetNames[0]]);
      data.forEach((item: any) => {
        let key = item.ClientBusinesssId + '-' + item.ReceiverCountry;
        if (summedData[key]) {
          summedData[key] += 1;
        } else {
          summedData[key] = 1;
        }
      });
      let result = Object.keys(summedData).map((key) => {
        let [ClientBusinesssId, ReceiverCountry] = key.split('-');
        return {
          ClientBusinesssId: parseInt(ClientBusinesssId),
          ReceiverCountry: ReceiverCountry,
          count: summedData[key],
        };
      });
      console.log(result);
      this.updatedExcelObj = {
        fileName: fileName,
        data: result,
        extension: extension,
        sheetName: sheetNames[0],
      };
    };
    event.target.value = null;
  }
  public downloadExcel(obj: any) {
    const worksheet = XLSX.utils.json_to_sheet(obj.data);

    // Create a workbook
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, obj.sheetName);

    // Convert workbook to binary Excel file
    const excelBuffer = XLSX.write(workbook, {
      bookType: obj.extension,
      type: 'array',
    });

    // Save the file with FileSaver.js
    const excelFile = new Blob([excelBuffer], {
      type: 'application/octet-stream',
    });
    saveAs(excelFile, `${obj.fileName}.${obj.extension}`);
  }
}
