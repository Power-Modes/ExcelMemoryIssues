import { Component } from "@angular/core";
import { v4 as uuidv4 } from 'uuid';
// images references in the manifest
import "../../../assets/icon-16.png";
import "../../../assets/icon-32.png";
import "../../../assets/icon-80.png";
const template = require("./app.component.html");
/* global console, Excel, require */

@Component({
	selector: "app-home",
	template
})
export default class AppComponent {
	welcomeMessage = "Welcome";

	// async run() {
	// 	try {
	// 		await Excel.run(async context => {
	// 			/**
	// 			 * Insert your Excel code here
	// 			 */
	// 			const range = context.workbook.getSelectedRange();

	// 			// Read the range address
	// 			range.load("address");

	// 			// Update the fill color
	// 			range.format.fill.color = "yellow";

	// 			await context.sync();
	// 			console.log(`The range address was ${range.address}.`);
	// 		});
	// 	} catch (error) {
	// 		console.error(error);
	// 	}
	// }

	private tableRowCount: number;
	private rangeRow: number;
	private enablePropertyBag: boolean = false;
	private rangeData: Array<any>;
	private tableData: Array<any>;

	dataString: string;

	constructor() {

		this.rangeData = new Array<any>();
		this.tableData = new Array<any>();

		this.createTable();
		this.createRange();
	}

	togglePropertyBag() {
		this.enablePropertyBag = !this.enablePropertyBag;
	}

	private async createTable() {
		try {
			await Excel.run(async context => {

				var sheet = context.workbook.worksheets.getActiveWorksheet();
				var table = sheet.tables.getItemOrNullObject("ItemTable");
				table.load(['rows/count']);
				await context.sync();

				if (table.isNullObject) {
					table = sheet.tables.add("A1:J1", true /*hasHeaders*/);
					table.name = "ItemTable";
					table.getHeaderRowRange().values = [["ID", "Row Id", "Order", "Item Name", "Item Type", "Start Date", "End Date", "Duration", "Progress", "Work"]];
					this.tableRowCount = 0;
				}
				else {
					this.tableRowCount = table.rows.count;
				}

				await context.sync();

				console.log('table row count', this.tableRowCount);
			});
		} catch (error) {
			console.error(error);
		}
	}

	private async createRange() {
		try {
			await Excel.run(async context => {

				var sheet = context.workbook.worksheets.getActiveWorksheet();
				var range = sheet.getRange('L1:U1');
				range.values = [["ID", "Row Id", "Order", "Item Name", "Item Type", "Start Date", "End Date", "Duration", "Progress", "Work"]];

				await context.sync();

				this.rangeRow = 2;

				console.log('range created');
			});
		} catch (error) {
			console.error(error);
		}
	}

	addTableItem() {
		try {
			Excel.run(async context => {

				var sheet = context.workbook.worksheets.getActiveWorksheet();
				var table = sheet.tables.getItemOrNullObject("ItemTable");

				let range: Excel.Range;

				if (this.tableRowCount == 0) {
					this.tableRowCount = 1;
					range = table.getDataBodyRange();
				}
				else {
					this.tableRowCount = this.tableRowCount + 1;
					range = table.getDataBodyRange().getRowsBelow(1);
				}

				range.load('address');
				await context.sync();

				var item = [this.tableRowCount, '=Row()', this.tableRowCount, 'Item ' + this.tableRowCount, 'Task', '09/25/2020', '09/26/2020', 1, 0, 0];

				range.values = [item];

				if (this.enablePropertyBag) {
					this.tableData.push(item);
					Office.context.document.settings.set('GanttData1', this.tableData);
					Office.context.document.settings.saveAsync();
				}

				await context.sync();

				console.log('new row range address', range.address);
			});
		} catch (error) {
			console.error(error);
		}
	}

	addRangeItem() {
		try {
			Excel.run(async context => {

				var sheet = context.workbook.worksheets.getActiveWorksheet();
				var range = sheet.getRange(`L${this.rangeRow}:U${this.rangeRow}`);

				var item = [this.rangeRow, '=Row()', this.rangeRow, 'Item ' + this.rangeRow, 'Task', '09/25/2020', '09/26/2020', 1, 0, 0];
				range.values = [item];

				if (this.enablePropertyBag) {
					this.rangeData.push(item);
					Office.context.document.settings.set('GanttData2', this.rangeData);
					Office.context.document.settings.saveAsync();
				}

				await context.sync();

				console.log('new row range done');
				this.rangeRow = this.rangeRow + 1;
			});
		} catch (error) {
			console.error(error);
		}
	}

	loadTableData() {
		var data = Office.context.document.settings.get('GanttData1');
		this.dataString = JSON.stringify(data);
	}

	loadRangeData() {
		var data = Office.context.document.settings.get('GanttData2');
		this.dataString = JSON.stringify(data);
	}
}
