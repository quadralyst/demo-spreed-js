import { Validators, FormBuilder, FormGroup } from '@angular/forms';
import * as GC from '@grapecity/spread-sheets';
import { jsonData } from '../app.data';
import '@grapecity/spread-sheets-charts';



import { Component, NgModule, enableProdMode, ViewChild, OnInit } from '@angular/core';
import { BrowserModule } from '@angular/platform-browser';
import { platformBrowserDynamic } from '@angular/platform-browser-dynamic';
import { SpreadSheetsModule } from '@grapecity/spread-sheets-angular';
import * as ExcelIO from "@grapecity/spread-excelio";
// import * as Excel from '@grapecity/spread-excelio';
import '@grapecity/spread-sheets-charts';
// import ExcelIO from "@grapecity/spread-excelio";
import { SpreadSheetsComponent } from '@grapecity/spread-sheets-angular/src/spreadSheets.component';
import { map } from 'rxjs';
import { style } from '@angular/animations';









@Component({
  selector: 'app-create-tamplate',
  templateUrl: './create-tamplate.component.html',
  styleUrls: ['./create-tamplate.component.scss']
})
export class CreateTamplateComponent implements OnInit {
  public createReportForm: FormGroup;
  spread: GC.Spread.Sheets.Workbook;
  importExcelFile: any;
  public isCopied = false;
  public sheetData: any;
  public showBackButton = false;
  public loadingBtn = false;
  hostStyle = {
    width: 'calc(100% - 280px)',
    height: '500px',
    overflow: 'hidden',
    float: 'left'
  };
  hostStylePreivew = {
    width: 'calc(100% - 280px)',
    height: '500px',
    overflow: 'hidden',
    float: 'left'
  };
  public showExelSheet = false;
  public previewExelSheet = false;
  public workbookObjData: any;
  public startRow :any;
  public templateData = [
    {
      name : "Report 1"
    },
    {
      name : "Report 2"
    },
    {
      name : "Report 3"
    },
    {
      name : "Report 4"
    },
    {
      name : "Report 5"
    },

  ]
  public headerJsonData = [
    {
      sNo: 1,
      firstName: "Gunjan",
      lastName: "Karun",
      department: "IT",
      loanAmount: 200000,
      payments: 100000,
    },
    {
      sNo: 2,
      firstName: "Umesh",
      lastName: "Carpenter",
      department: "Menufecharing",
      loanAmount: 300000,
      payments: 400000,
    },
    {
      sNo: 3,
      firstName: "Devendra",
      lastName: "Mhaski",
      department: "Development",
      loanAmount: 500000,
      payments: 150000,
    },
    {
      sNo: 4,
      firstName: "Retesh",
      lastName: "sahu",
      department: "IT",
      loanAmount: 600000,
      payments: 10000,
    },
  ]
  public dataKeys: any;
  public dataValues: any;
  startForm: any;
  public selectRow : any;
  public selectCol : any;
  constructor(
    private fb: FormBuilder,
  ) { }

  ngOnInit(): void {
    this.createReportForm = this.fb.group({
      reportName: ['', Validators.compose([Validators.required])],
      importExcelFiles: ['', Validators.compose([Validators.required])]
    });
    this.getData();
  }

  public getData() {
    this.dataKeys = Object.keys(this.headerJsonData[0]);
  }

  changeFileDemo(e: any) {
    console.log('eeeeeeeeeeee', e)
    this.importExcelFile = e.target.files[0];
  }
public setStartRow(e:any) {
  console.log("startRow=>",this.startRow)
}
  public CancelLoadExcel() {
    this.sheetData = '';
    this.importExcelFile = '';
    this.showExelSheet = false;
    this.createReportForm.get('importExcelFiles') ?.reset();
    this.createReportForm.reset();
  }

  loadExcel(e: any) {
    if (this.createReportForm.value.importExcelFiles) {
      this.showBackButton = false;
      this.showExelSheet = true;
      this.previewExelSheet = false;
      console.log('this.importExcelFile', this.importExcelFile)
      let spread = this.spread;
      let excelIo = new ExcelIO.IO();
      // this.excelIO = new Excel.IO();
      let excelFile = this.importExcelFile;
      // let password = this.password;
      let incrementalEle = document.getElementById("incremental") as HTMLInputElement;
      let loadingStatus = document.getElementById("loadingStatus") as HTMLInputElement;

      console.log('excelFile=>', excelFile);
      console.log('excelIo=>', excelIo);

      // here is excel IO API
      excelIo.open(excelFile, (json: any) => {
        let workbookObj = json;
        // console.log('workbookObj*********************', workbookObj);




        // if (incrementalEle.checked) {
        //   spread.fromJSON(workbookObj, {
        //     incrementalLoading: {
        //       loading: function (progress: any, args: any) {
        //         progress = progress * 100;
        //         loadingStatus.value = progress;
        //         console.log("current loading sheet", args.sheet && args.sheet.name());
        //       },
        //       loaded: function () {
        //       }
        //     }
        //   });
        // } else {
        // }
        spread.fromJSON(workbookObj);
        this.workbookObjData = workbookObj;





        // let sheet = this.spread.getActiveSheet();
        // let slectedRow = 3;
        // this.headerJsonData.map(res => {
        //   let newRow = slectedRow + 1
        //   sheet.addRows(newRow, 1);
        //   sheet.copyTo(3, 0, newRow, 0, 1, 7, GC.Spread.Sheets.CopyToOptions.value);
        //   sheet.copyTo(3, 0, newRow, 0, 1, 7, GC.Spread.Sheets.CopyToOptions.style);
        //   sheet.copyTo(3, 0, newRow, 0, 1, 7, GC.Spread.Sheets.CopyToOptions.formula);
        //   slectedRow = slectedRow + 1;
        // })


        // let seetName: string = '';
        // spread.fromJSON(json, {
        //   incrementalLoading: {
        //     loading: function (progress: any, args: any) {
        //       seetName = args.sheet && args.sheet.name();
        //       alert(seetName)
        //     },
        //     loaded: function () {
        //     }
        //   }
        // });
        // const dataTable = json.sheets[seetName].data.dataTable;
        // const arr: any = [];
        // Object.keys(dataTable[3]).map((val: any) => {
        //   console.log('val****************', val)
        //   // const vals = Object.values(dataTable[newRow][val]);
        //   // const obj = { index: val, value: vals[0], style: vals[1], formula: vals[2] }
        //   // arr.push(obj);
        // })


   




        // sheet.addRows(5, 1);
        // sheet.copyTo(3, 0, 5, 0, 1, 7, GC.Spread.Sheets.CopyToOptions.value);
        // sheet.copyTo(3, 0, 5, 0, 1, 7, GC.Spread.Sheets.CopyToOptions.style);
        // sheet.copyTo(3, 0, 5, 0, 1, 7, GC.Spread.Sheets.CopyToOptions.formula);



      }, function (e: any) {
        console.log('e.errorMessage', e.errorMessage)
        // process error
        alert(e.errorMessage + '=============');
      },
        // { password: password }
      );
    } else {
      alert('Please select file')
    }
  }


  public backToSheet() {
    this.showBackButton = false;
    let spread = this.spread;
    let excelIo = new ExcelIO.IO();
    let json: any = spread.toJSON();
    console.log('sheetData', this.sheetData)
    // SHEET START ROW DATA
    spread.fromJSON(this.sheetData);

  }



  preview(e:any){
    this.loadingBtn = true;
    // alert()
    console.log('loadingBtn=>',this.loadingBtn)
    
    let spread = this.spread;
    let json: any = spread.toJSON();
    let sheet = this.spread.getActiveSheet();
    // let startFrom = 3;
    let startFrom = this.startRow -1;
    let newRow = startFrom + 1;
    let arr: any[] = [];
    const dataTable = json.sheets[sheet.name()].data.dataTable;
    Object.keys(dataTable[newRow]).map((val: any) => {
      const vals = Object.values(dataTable[newRow][val]);
      const obj = { index: val, value: vals[0], style: vals[1], formula: vals[2] }
      arr.push(obj);
    })

    this.headerJsonData.map((res:any ,index :any) => {
      console.log('************res******', res);
      sheet.addRows(newRow, 1);
      sheet.copyTo(startFrom, 0, newRow, 0, 1, arr.length, GC.Spread.Sheets.CopyToOptions.value);
      sheet.copyTo(startFrom, 0, newRow, 0, 1, arr.length, GC.Spread.Sheets.CopyToOptions.style);
      sheet.copyTo(startFrom, 0, newRow, 0, 1, arr.length, GC.Spread.Sheets.CopyToOptions.formula);
      const cellArr = Array.from(Array(arr.length).keys())
      cellArr.map(resonse => {
        const cellVal = sheet.getValue(newRow, resonse);
        sheet.setValue(newRow, resonse, res[cellVal], GC.Spread.Sheets.SheetArea.viewport);
      })
      newRow = newRow + 1 
      // newRow = slectedRow + 1;
      if(this.headerJsonData.length -1 === index) {
        // alert(1)
       this.loadingBtn = false;
      }
    })
    sheet.deleteRows(3,1);
   
  }

  initSpread($event: any) {
    console.log('event============', $event)

    this.spread = $event.spread;
    let spread = this.spread;
    spread.options.calcOnDemand = true;
  }


  initSpreadpreview($event: any) {
    console.log('event============', $event)

    this.spread = $event.spread;
    let spread = this.spread;
    spread.options.calcOnDemand = true;
  }



  public clickToCopyText(el: any): void {
    console.log('this.workbookObjData#######################', this.workbookObjData)
    this.isCopied = true;
    if(el) {
      console.log(el);
      if(el == 'sNo') {
        console.log('start form ', this.selectRow);
      }
      this.spread.sheets[1].setValue(this.selectRow, this.selectCol, el);
    }

    navigator.clipboard.writeText(el);
    setTimeout(() => {
      this.isCopied = false;
    }, 2000)
  }

public repeatRow(type :any) {
  console.log('type ==>',type);
}

  public getCellValue(e: any) {
    // alert()
    // return
    console.log('clieck cell event get row', e.row)
    console.log('clieck cell event get col', e.col)
    this.selectRow = e.row;
    this.selectCol = e.col;
    console.log('getValue', this.spread.sheets[1].getValue(this.spread.sheets[1].getActiveRowIndex(), this.spread.sheets[1].getActiveColumnIndex()));
    console.log('get formula', this.spread.sheets[1].getFormula(this.spread.sheets[1].getActiveRowIndex(), this.spread.sheets[1].getActiveColumnIndex()));
  }



  public editChange(e: any) {
    console.log('edit change', e)
    if (e.editingText && e.editingText == "repeat") {
      this.startForm = e.row;
      console.log('startForm=>', this.startForm)
    }
  }

  public clipboardPasted(e: any) {
    console.log('clipboardPasted', e)
    console.log('clipboardPasted e.pasteData.text', e.pasteData.text)
    if (e.pasteData && e.pasteData.text == "repeat") {
      this.startForm = e.row;
      console.log('startForm=>', this.startForm)
    }
  }

  public userFormulaEntered(e: any) {
    console.log('userFormulaEntered', e);
  }
  
  public editEnd(e: any) {
    console.log('editEnd', e);
  }
  public rangeSorting(e: any) {
    console.log('rangeSorting', e);
  }

}
