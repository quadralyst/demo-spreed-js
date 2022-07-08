
import { Component, NgModule, enableProdMode, ViewChild } from '@angular/core';
import { BrowserModule } from '@angular/platform-browser';
import { platformBrowserDynamic } from '@angular/platform-browser-dynamic';
import { SpreadSheetsModule } from '@grapecity/spread-sheets-angular';
import * as GC from '@grapecity/spread-sheets';
import * as ExcelIO from "@grapecity/spread-excelio";
// import * as Excel from '@grapecity/spread-excelio';
import '@grapecity/spread-sheets-charts';
// import ExcelIO from "@grapecity/spread-excelio";
import { jsonData } from './app.data';
import { SpreadSheetsComponent } from '@grapecity/spread-sheets-angular/src/spreadSheets.component';
import { map } from 'rxjs';
import { style } from '@angular/animations';
// import './styles.css';
declare var saveAs: any;
// window.GC = GC;
import { Validators, FormBuilder, FormGroup } from '@angular/forms';

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.scss']
})
export class AppComponent {
  title = 'sjs-angular-app';
  spread: GC.Spread.Sheets.Workbook;


  hostStyle = {
    width: 'calc(100% - 280px)',
    height: '500px',
    overflow: 'hidden',
    float: 'left'
};
importExcelFile: any;
exportFileName = "export.xlsx";
password: string;
public chack :any ; 
public createReportForm: FormGroup;
 

  constructor( 
    private fb: FormBuilder,

  ) {
    this.createReportForm = this.fb.group({
        reportName: ['', Validators.compose([Validators.required])],
        importExcelFiles: ['', Validators.compose([Validators.required])]
      });
   }


  initSpread($event: any) {

      this.spread = $event.spread;
    let spread = this.spread;
    spread.options.calcOnDemand = true;
    
    const proData = [
        {
            number : 1,
            name : 'Umesh carpenter',
            price : 10,
        },
        {
            number : 2,
            name : 'swapan',
            price : 20,
        },
        {
            number : 3,
            name : 'Vishnu',
            price : 30,
        },
        {
            number : 4,
            name : 'Prateek',
            price : 40,
        },
        {
            number : 5,
            name : 'Devendra ',
            price : 50,
        },
        {
            number : 6,
            name : 'Ram',
            price : 60,
        },
        {
            number : 7,
            name : 'James',
            price : 70,
        },
        {
            number : 8,
            name : 'Gagan',
            price : 80,
        },
        {
            number : 9,
            name : 'Sumit',
            price : 90,
        },
        {
            number : 10,
            name : 'Amit',
            price : 100,
        },
        {
            number : 11,
            name : 'Keshaw',
            price : 110,
        }
    ];
    const dataTable = jsonData.sheets['Normal Acc'].data.dataTable;
    // const dataRow:any = jsonData.sheets['Normal Acc'].data.dataRow;
    const dataRow:any = {};

    let startForm = 4;
    proData .map(res =>{
        let key = String(startForm);
        dataRow[key] = {'0': {value : res.number, style : "__builtInStyle25"},'1': {value : '', style : "__builtInStyle26"},'2': {value : '', style : "__builtInStyle27"},'3': {value : res.name, style : "__builtInStyle28"},'4': {value : '', style : "__builtInStyle26"},'5': {value : '', style : "__builtInStyle26"},'6': {value : res.price, style : "__builtInStyle29"},'7': {value : "", style : "__builtInStyle30"},'8': {value : 0, style: "__builtInStyle31" , },'9': {value : '', style : "__builtInStyle29"},'10': {value : '', style : "__builtInStyle29"},'11': {value : '', style : "__builtInStyle29"},'12': {value : '', style : "__builtInStyle29"},'13': {value : '', style : "__builtInStyle29"},'14': {value : '', style : "__builtInStyle29"},'15': {value : '', style : "__builtInStyle29"}}        
        startForm = startForm + 1;        
        Object.assign(dataTable, dataRow);
        
        // let sheet:any = spread.getActiveSheet();
      
        // console.log('indexd==>',index)

    // sheet.setValue(startForm,2,res.name, GC.Spread.Sheets.SheetArea.viewport);


    })
    // let footerData:any = {     

    //     "0": {
    //         "style": "__builtInStyle46"
    //         },
    //         "1": {
    //             "style": "__builtInStyle46",
    //         },
    //         "2": {
    //           "style": "__builtInStyle46"
    //         },
    //         "3": {
    //           "style": "__builtInStyle46"
    //         },
    //         "4": {
    //           "style": "__builtInStyle46"
    //         },
    //         "5": {
    //           "style": "__builtInStyle46"
    //         },
    //         "6": {
    //           "style": "__builtInStyle46",
    //           "value": 0,
    //           "formula": "SUM(G4:G10)"
    //         },
    //         "7": {
    //           "style": "__builtInStyle46",
    //           "value": 0,
    //           "formula": "SUM(H4:H10)"
    //         },
    //         "8": {
    //           "style": "__builtInStyle46",
    //           "value": {
    //             "_error": "#VALUE!",
    //             "_code": 15
    //           },
    //           "formula": "SUM(I4:I10)"
    //         },
    //         "9": {
    //           "style": "__builtInStyle46",
    //           "value": 0,
    //           "formula": "SUM(J4:J10)"
    //         },
    //         "10": {
    //           "style": "__builtInStyle46",
    //           "value": 0,
    //           "formula": "SUM(K4:K10)"
    //         },
    //         "11": {
    //           "style": "__builtInStyle46",
    //           "value": 0,
    //           "formula": "SUM(L4:L10)"
    //         },
    //         "12": {
    //           "style": "__builtInStyle46",
    //           "value": 0,
    //           "formula": "SUM(M4:M10)"
    //         },
    //         "13": {
    //           "style": "__builtInStyle46",
    //           "value": 0,
    //           "formula": "SUM(N4:N10)"
    //         },
    //         "14": {
    //           "style": "__builtInStyle46",
    //           "value": 0,
    //           "formula": "SUM(O4:O10)"
    //         },
    //         "15": {
    //           "style": "__builtInStyle46",
    //           "value": 0,
    //           "formula": "SUM(P4:P10)"
    //         }
    //       }    
    const footerData:any = {};
    
    const foterCount = 3 + proData.length + 1;
    // const total = Number(ss) - 1;
    const footer:any = {};
    footer[foterCount]= footerData;
    footer[foterCount][6].formula = 'SUM(G4:G'+ foterCount+')';
    footer[foterCount][7].formula = 'SUM(H4:H'+ foterCount+')';
    footer[foterCount][9].formula = 'SUM(J4:J'+ foterCount+')';
    footer[foterCount][10].formula = 'SUM(K4:K'+ foterCount+')';
    footer[foterCount][11].formula = 'SUM(L4:L'+ foterCount+')';
    footer[foterCount][12].formula = 'SUM(M4:M'+ foterCount+')';
    footer[foterCount][13].formula = 'SUM(N4:N'+ foterCount+')';
    footer[foterCount][14].formula = 'SUM(O4:O'+ foterCount+')';
    footer[foterCount][15].formula = 'SUM(P4:P'+ foterCount+')';
    // Object.assign(dev, dd);
    // Object.assign(dataTable, footer);        

    console.log('dataRow', dataRow);

    
    console.log('dev*********************', dataTable);


spread.fromJSON(jsonData); //Use fromJSON method to load the template, here template variable field is defined in binding.js.
let sheet = spread.getActiveSheet();
//  sheet.setDataSource(new GC.Spread.Sheets.Bindings.CellBindingSource(dataRow).setValue); 

// sheet.addRows(0,2);
// sheet.setValue(5,2,"ColumnHeader", GC.Spread.Sheets.SheetArea.viewport);
 
    
}

changeFileDemo(e: any) {
    this.importExcelFile = e.target.files[0];
}
changePassword(e: any) {
    this.password = e.target.value;
}
changeExportFileName(e: any) {
  console.log('444444444=>',e.target.value);
    this.exportFileName = e.target.value;
}
changeIncremental(e: any) {
    this.chack = e.target.checked ? true : false;
}
loadExcel(e: any) {
    console.log('this.importExcelFile', this.importExcelFile)
    let spread = this.spread;
    let excelIo = new ExcelIO.IO();
    // this.excelIO = new Excel.IO();
    let excelFile = this.importExcelFile;
    let password = this.password;
    let incrementalEle = document.getElementById("incremental") as HTMLInputElement;
    let loadingStatus = document.getElementById("loadingStatus") as HTMLInputElement;

    console.log('excelFile=>',excelFile);
    console.log('excelIo=>',excelIo);

    // here is excel IO API
    excelIo.open(excelFile, function (json: any) {
        let workbookObj = json;


        const impoetData = {
            "7": {
                "0": {
                    "style": "__builtInStyle25",
                    "value": 4
                },
                "1": {
                    "style": "__builtInStyle26",
                    "value": "<C_ID>"
                },
                "2": {
                    "style": "__builtInStyle27"
                },
                "3": {
                    "style": "__builtInStyle28",
                    "value": "<C_Name>"
                },
                "4": {
                    "style": "__builtInStyle26"
                },
                "5": {
                    "style": "__builtInStyle26"
                },
                "6": {
                    "style": "__builtInStyle29"
                },
                "7": {
                    "style": "__builtInStyle30"
                },
                "8": {
                    "style": "__builtInStyle31",
                    "value": 0,
                    "formula": "G5+H5"
                },
                "9": {
                    "style": "__builtInStyle32"
                },
                "10": {
                    "style": "__builtInStyle29"
                },
                "11": {
                    "style": "__builtInStyle33"
                },
                "12": {
                    "style": "__builtInStyle34"
                },
                "13": {
                    "style": "__builtInStyle35"
                },
                "14": {
                    "style": "__builtInStyle36"
                },
                "15": {
                    "style": "__builtInStyle37"
                }
            },
            "8": {
                "0": {
                    "style": "__builtInStyle25",
                    "value": 5
                },
                "1": {
                    "style": "__builtInStyle26",
                    "value": "<C_ID>"
                },
                "2": {
                    "style": "__builtInStyle27"
                },
                "3": {
                    "style": "__builtInStyle28",
                    "value": "<C_Name>"
                },
                "4": {
                    "style": "__builtInStyle26"
                },
                "5": {
                    "style": "__builtInStyle26"
                },
                "6": {
                    "style": "__builtInStyle29"
                },
                "7": {
                    "style": "__builtInStyle30"
                },
                "8": {
                    "style": "__builtInStyle31",
                    "value": 0,
                    "formula": "G5+H5"
                },
                "9": {
                    "style": "__builtInStyle32"
                },
                "10": {
                    "style": "__builtInStyle29"
                },
                "11": {
                    "style": "__builtInStyle33"
                },
                "12": {
                    "style": "__builtInStyle34"
                },
                "13": {
                    "style": "__builtInStyle35"
                },
                "14": {
                    "style": "__builtInStyle36"
                },
                "15": {
                    "style": "__builtInStyle37"
                }
            },
            "9": {
                "0": {
                    "style": "__builtInStyle25",
                    "value": 6
                },
                "1": {
                    "style": "__builtInStyle26",
                    "value": "<C_ID>"
                },
                "2": {
                    "style": "__builtInStyle27"
                },
                "3": {
                    "style": "__builtInStyle28",
                    "value": "<C_Name>"
                },
                "4": {
                    "style": "__builtInStyle26"
                },
                "5": {
                    "style": "__builtInStyle26"
                },
                "6": {
                    "style": "__builtInStyle29"
                },
                "7": {
                    "style": "__builtInStyle30"
                },
                "8": {
                    "style": "__builtInStyle31",
                    "value": 0,
                    "formula": "G5+H5"
                },
                "9": {
                    "style": "__builtInStyle32"
                },
                "10": {
                    "style": "__builtInStyle29"
                },
                "11": {
                    "style": "__builtInStyle33"
                },
                "12": {
                    "style": "__builtInStyle34"
                },
                "13": {
                    "style": "__builtInStyle35"
                },
                "14": {
                    "style": "__builtInStyle36"
                },
                "15": {
                    "style": "__builtInStyle37"
                }
            }
        }



        const inportTable = workbookObj.sheets['Normal Acc'].data.dataTable;

        console.log('dev*********************', inportTable);
        
        Object.assign(inportTable, impoetData);
        
        console.log('workbookObj*********************', workbookObj);




        if (incrementalEle.checked) {
            spread.fromJSON(workbookObj, {
                incrementalLoading: {
                    loading: function (progress: any, args: any) {
                        progress = progress * 100;
                        loadingStatus.value = progress;
                        console.log("current loading sheet", args.sheet && args.sheet.name());
                    },
                    loaded: function () {
                    }
                }
            });
        } else {
            spread.fromJSON(workbookObj);
        }
    }, function (e: any) {
        // process error
        alert(e.errorMessage);
    }, { password: password });
}
saveExcel(e: any) {
  console.log('123',e.target.value)
    let spread = this.spread;
    let excelIo = new ExcelIO.IO();

    let fileName = this.exportFileName;
    let password = this.password;
    if (fileName.substring(-5, 5) !== '.xlsx') {
        fileName += '.xlsx';
    }

    let json = spread.toJSON();
    console.log('json==>',json)
    console.log('fileName==>',fileName)
    // here is excel IO API
    excelIo.save(json, function (blob: any) {
        saveAs(blob, fileName);
    }, function (e: any) {
        // process error
        console.log('umesh',e)
        console.log(e);
    }, { password: password });

}


workbookInit(args:any) {
    let spread = new GC.Spread.Sheets.Workbook(document.getElementById("spreadSheet") as HTMLInputElement, { sheetCount: 3 });
    // spread.fromJSON(jsonData);
    console.log('spread==>',spread);
    
    



//   let spread: GC.Spread.Sheets.Workbook = args.spread;
//   let index = spread.getActiveSheetIndex() 
//     alert(index)
//   let sheet = spread.getActiveSheet();
//   sheet.autoGenerateColumns = true;
//   sheet.getCell(0, 0).text("My SpreadJS Angular Project").foreColor("blue");
  
}
}



