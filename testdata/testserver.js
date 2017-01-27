var app = require("express")();

var test = require(__dirname + "/main.js");
var test = new test(__dirname + '/testdata/sample.xls');

app.get('/',function(req,res){
	test.setToEndpoint(res);
})
app.listen(3000);