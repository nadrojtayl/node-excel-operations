"use strict";

var xlsx = require('node-xlsx');
var fs = require("fs");
class XLShelper {
	constructor(directory){
		this.data = require('node-xlsx').parse(directory)[0].data; 
		this.columnNames = this.data[0];
		this.reusableColumnObj = {};
		var that = this;
		this.data[0].forEach(function(name,ind){
			that.columnNames[name] = ind;
			that.reusableColumnObj[name];
		})
		this.pivotTable;
	}

	write(directory){
		this.buffer = require('node-xlsx').build([{name: "sheetName", data: this.data}]);

		fs.writeFile(directory,this.buffer,function(err){
			console.log(err);
		})
	}

	seeColumnNames(){
		console.log(this.data[0])
		return this.data[0];
	}

	addColumn(colName,cb){
		this.data[0].push(colName);
		var that = this;
		for(var i = 1;i< this.data.length;i++){
			var row = this.data[i];
			this.columnNames.forEach(function(val,ind){
				that.reusableColumnObj[val] = row[ind];
			})
			try {
				row.push(cb(that.reusableColumnObj));
			} catch(err){
				console.log("The column names you entered probably didn't match your columns")	
			}
		}

	}

	addColumnSimple(colName,calcStr){
		var that = this;
		this.data[0].push(colName);
		for(var i = 1;i< this.data.length;i++){
			var row = this.data[i];
			this.columnNames.forEach(function(name,ind){
				global[name] = row[ind];
			})
			try {
				row.push(eval(calcStr));
			} catch(err){
				console.log("The column names you entered probably didn't match your columns")	
			}
		}
	}

	addRow(){

	}

	pivotTable(fnction,colToOperateOn,rowlabel,columnlabel){
		
		var table = [];
		var rowInd = this.columnNames[rowlabel];
		var colInd = this.columnNames[columnlabel];

		var rowNameMap = {};
		var colNameMap = {};

		var Values = {};

		var OpIndex = this.columnNames[colToOperateOn]
		var row;
		var toprow = [""];

		//get all rows and columns
		for(var i = 1;i<this.data.length;i++){
			row = this.data[i]
			if(rowNameMap[row[rowInd]] === undefined){
				rowNameMap[row[rowInd]] = true;
			} 
			if(colNameMap[row[colInd]] === undefined){
				colNameMap[row[colInd]] = true;
			} 
		}
		
		//	create matrix of Values
		for(var rName in rowNameMap){
			var newObj = {};
			for(var cName in colNameMap){
				newObj[cName] = [];
			}
			Values[rName] = newObj;
		}

		for(var i =1;i<this.data.length;i++){
			var row = this.data[i];
			Values[row[rowInd]][row[colInd]].push(row[OpIndex]);
		}



		for(var cKey in colNameMap){
			toprow.push(cKey)
		}

		table.push(toprow);

		

		for(var key in Values){
			var newrow = [key];
			for(var cKey in Values[key]){
				if(Values[key][cKey].length === 0){
					newrow.push(0);
				} else {
					newrow.push(Values[key][cKey].reduce(fnction));
				}
			}
			table.push(newrow)
		}



		this.pivotTable = table;

	}

	writepivotTable(directory){
		this.buffer = require('node-xlsx').build([{name: "sheetName", data: this.pivotTable}]);
		fs.writeFile(directory,this.buffer,function(err){
			console.log(err);
		})

	}

	printToHTML(){

	}

	setToEndpoint(){

	}

}



var test = new XLShelper(__dirname + '/sample.xls');
// test.addColumnSimple("BodyWeight+BrainWeight","bodywt + brainwt")
// // test.write(__dirname + "/test.xls");
test.pivotTable(function(a,b){return a+b},"bodywt","age","sex#")
test.writepivotTable(__dirname + "/test.xls")