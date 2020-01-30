import { Component } from "@angular/core";
import { HttpClient } from "@angular/common/http";

@Component({
  selector: "app-root",
  templateUrl: "./app.component.html",
  styleUrls: ["./app.component.scss"]
})
export class AppComponent {
  welcomeMessage = "Excel is running in Angular!";

  constructor(private http: HttpClient) {}

  async run() {
    try {
      await Excel.run(async context => {
        /**
         * Insert your Excel code here
         */
        const range = context.workbook.getSelectedRange();

        // Read the range address
        range.load("address");

        // Update the fill color
        range.format.fill.color = "yellow";

        this.http
          .get<any>("https://jsonplaceholder.typicode.com/albums")
          .subscribe(arg => {
            console.log("data", arg);
          });

        await context.sync();
        console.log(`The range address was ${range.address}.`);
      });
    } catch (error) {
      console.error(error);
    }
  }

  async createTable() {
    try {
      await Excel.run(async context => {
        /**
         * Insert your Excel code here
         */
        const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
        const expensesTable = currentWorksheet.tables.add(
          "A1:D1",
          true /*hasHeaders*/
        );
        expensesTable.name = "ExpensesTable";

        expensesTable.getHeaderRowRange().values = [
          ["Date", "Merchant", "Category", "Amount"]
        ];

        expensesTable.rows.add(null /*add at the end*/, [
          ["1/1/2017", "The Phone Company", "Communications", "120"],
          ["1/2/2017", "Northwind Electric Cars", "Transportation", "142.33"],
          ["1/5/2017", "Best For You Organics Company", "Groceries", "27.9"],
          ["1/10/2017", "Coho Vineyard", "Restaurant", "33"],
          ["1/11/2017", "Bellows College", "Education", "350.1"],
          ["1/15/2017", "Trey Research", "Other", "135"],
          ["1/15/2017", "Best For You Organics Company", "Groceries", "97.88"]
        ]);

        expensesTable.columns.getItemAt(3).getRange().numberFormat = [
          ["€#,##0.00"]
        ];
        expensesTable.getRange().format.autofitColumns();
        expensesTable.getRange().format.autofitRows();

        await context.sync();
        // console.log(`The range address was ${range.address}.`);
        // console.log("test Latino");
      });
    } catch (error) {
      console.error(error);
    }
  }

  async createChart() {
    try {
      await Excel.run(async context => {
        /**
         * Insert your Excel code here
         */

        const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
        const expensesTable = currentWorksheet.tables.getItem("ExpensesTable");
        const dataRange = expensesTable.getDataBodyRange();

        const chart = currentWorksheet.charts.add(
          Excel.ChartType.columnClustered,
          dataRange
        );

        chart.setPosition("A15", "F30");
        chart.title.text = "Expenses";
        // chart.legend.position = "right";
        chart.legend.format.fill.setSolidColor("white");
        chart.dataLabels.format.font.size = 15;
        chart.dataLabels.format.font.color = "black";
        chart.series.getItemAt(0).name = "Value in €";

        await context.sync();
        // console.log(`The range address was ${range.address}.`);
        // console.log("test Latino");
      });
    } catch (error) {
      console.error(error);
    }
  }
}
