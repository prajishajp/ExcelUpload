import { Component } from '@angular/core';

import * as XLSX from 'xlsx';
import { Http, Response, RequestOptions, Headers } from '@angular/http';
import { transaction } from './transaction';
import {TransService} from './trans.service';
import { HttpClient, HttpHeaders} from '@angular/common/http';

type AOA = any[][];

@Component({
	selector: 'app-root',
	template: `
	<input type="file" (change)="onFileChange($event)" multiple="false"  accept="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" />
	<table class="sjs-table" border="striped">
		<tr *ngFor="let row of data">
			<td *ngFor="let val of row">
				{{val}}
			</td>
		</tr>
	</table>
	<button (click)="export()">Export</button>
	<button (click)="import()">Import</button>
	{{retPostData}}
	`
})

export class AppComponent {
	public retPostData
	res: any;
	constructor(private http:Http, private transService: TransService) {}
	httpService: any;
	data: AOA = [];
	
	transList :Array<transaction> = new Array<transaction>();
	
	wopts: XLSX.WritingOptions = { bookType: 'xlsx', type: 'array' };
	fileName: string = 'TRALive.xlsx';
	myFiles:string [] = [];
	sMsg:string = '';

	onFileChange(evt: any) {
		/* wire up file reader */
		const target: DataTransfer = <DataTransfer>(evt.target);
		if (target.files.length !== 1) throw new Error('Cannot use multiple files');
		const reader: FileReader = new FileReader();
		reader.onload = (e: any) => {
			/* read workbook */
			const bstr: string = e.target.result;
			const wb: XLSX.WorkBook = XLSX.read(bstr, {type: 'binary'});


			/* grab first sheet */
			const wsname: string = wb.SheetNames[0];
			const ws: XLSX.WorkSheet = wb.Sheets[wsname];

			/* save data */
			this.data = <AOA>(XLSX.utils.sheet_to_json(ws, {header: 1}));
			console.log(this.data);
			
			var transObj;
			for(var i=1;i<this.data.length;i++)
			{
				transObj =new transaction();
				console.log(this.data[i]);
				transObj.item=this.data[i][0]
				transObj.itemprofile=this.data[i][1]
				transObj.dealnumber=this.data[i][2]
				transObj.owner=this.data[i][3]
				transObj.date=this.data[i][4]
				transObj.time=this.data[i][5]
				transObj.statuss=this.data[i][6]
				transObj.currency=this.data[i][7]
				transObj.amount=this.data[i][8]
				transObj.localequivalent=this.data[i][9]
				transObj.valuedate=this.data[i][10]
				transObj.stepid=this.data[i][11]
				transObj.sender=this.data[i][12]
				transObj.customername=this.data[i][13]
				transObj.response=this.data[i][14]
				transObj.entity=this.data[i][15]
				transObj.department=this.data[i][16]
				transObj.section=this.data[i][17]
				transObj.exceededdate=this.data[i][18]
				
				this.transList.push(transObj);	
				console.log(this.transList);


			}

			this.transService.SaveTransaction(this.transList).subscribe
		(data=> {
			console.log(data);
		})



		};
		reader.readAsBinaryString(target.files[0]);
		
	}

	export(): void {
		/* generate worksheet */
		const ws: XLSX.WorkSheet = XLSX.utils.aoa_to_sheet(this.data);

		/* generate workbook and add the worksheet */
		const wb: XLSX.WorkBook = XLSX.utils.book_new();
		XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');

		/* save to file */
		XLSX.writeFile(wb, this.fileName);
	}
	
}