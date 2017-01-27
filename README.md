Built on top of the excellent work of mgcrea and node-xlsx (https://www.npmjs.com/package/node-xlsx)

Note: This library only handles single sheet XLS files - put all your data in one sheet before using this if you don't want to get an error

To use this library, you must have your column labels defined in the first row of your xls file

To install the library
```js
	var xlsHelper = require("node-excel-operations");
```

To read your file-- and create a helper to perform operations on your file -- pass in absolute directory path to a new instance of the library
```js
	var xls = new xlsHelper(__dirname + "/sampleData.xlsx");

````

Helper functions available to you:

write(filepath): writes the file (and any modifications you made) to a new xls file at provided filepath

```js
	xls.write(__dirname + "/newfile.xls") // this will produce a new xls file, if you made any changes to the file using functions below, the changes are reflected on the written file
```

seeColumnNames(): returns array of columnNames(defined as elements in first row in xls), also logs columnNames to the console

addColumn(columnName,callback) : adds a new Column to the end of the sheet under the name given in columnName. The callback defines the values present in the new column. The only argument passed to the callback is an object with all of your column names as keys. Return the mathematical result you want to put into your new column in your callback.

For example, if you have a column called "oldweight" and a column called "newweight", here is how you would add a column "old weight-new weight", which would have the difference of the neweight and oldweight columns, to your sheet

```js
xls.addColumn("Old Weight - New Weight Column",function(data){
	return data.oldweight - data.newweight;
}) 

```

To add the sum of the two columns you would do this:

```js
xls.addColumn("Old Weight + New Weight Column",function(data){
	return data.oldweight + data.newweight;
}) 

```

addColumnSimple(columnName,calcString): same as addColumn above, but simpler. The first argument is the name that will be given to your new column. The second argument is a string describing the mathematical result you want to put in your new column. Example:

```js
xls.addColumnSimple("Old Weight + New Weight Column","oldweight + newweight");

```

addRow(rowName,cb): adds a new row to the bottom of your sheet. The row will be given a label based on waht you pass in as row name. Similar to addColumn,above, addRow takes a callback that defines how you want to combine the elements in each column to get your new row. To get the sum of each column, you would do this:

```js
xls.addRow("Sum",function(a,b){return a + b;});

```

pivotTable(function,colToOperateOn,rowlabel,columnlabel): mimics Excel pivot tables (http://www.excel-easy.com/data-analysis/pivot-tables.html) 

The first argument is the function that defines what operation you want to do on the value in each row that you are "pivoting". The second is the name of the column you want to pivot. The third is the name of the column whose values you want to put in the row of your table. The fourth is the name of the row whose values you want to put in the columns of your table

//finishing soon

only works on numbers