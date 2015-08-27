var should = require('chai').should(),
    qube = require('../index'),
    excel = require('../lib/js-xlsx-templater'),
    workbook = excel.Workbook,
    worksheet = excel.Worksheet,   
    mergebooks = qube.MergeBooks;

//{"header":true,skipEmptyRows:true,"headers":["Goodbye","Bob"],startCell:"B6",colWidth:10}

describe('#Excel templatey tests', function(){
	
	it('Worksheets have correct range A1:B3', function(){
		var ws = worksheet(['A','B'],[[1,2],[3,4]],1,{});
		var range = "A1:B3" 
		ws['!ref'].should.equal(range);
	});

	it('Worksheets have correct range A1:B5', function(){
		var ws = worksheet(['A','B'],[[1,2],[3,4],[5,6],[7,8]],1,{});
		var range = "A1:B5" 
		ws['!ref'].should.equal(range);
	});

	it('Workbook should have 3 sheets', function(){
		var ws = worksheet(['A','B'],[[1,2]],1,{});
		var wb = new workbook({"One":ws,"Two":ws, "Three":ws});		
		wb.sheet_names().length.should.equal(3);
	});

	var ws = worksheet(['A','B'],[[1,2]],1,{});
	var template = new workbook({"One":ws,"Two":ws});
	var merged = mergebooks({"One":worksheet(['A','B'],[[1,2],[3,4],[5,6]],1,{})}, template.toFile());
	var wb = workbook(merged);

	it('Workbook sheet range should expand', function(){		
		var dataSheet = wb.read_sheet(wb.find_sheet_path("One"));
		var range = wb.sheet_datarange(dataSheet);
		range.should.equal("A1:B4");
	});

	it('Workbook should have 2 sheets', function(){		
		wb.sheet_names().length.should.equal(2);
	});

	it('Workbook can have sheets with spaces in their names', function(){
		var ws = worksheet(['A','B'],[[1,2]],1,{});
		var wb = new workbook({"O n e":ws,"T w o":ws, "Thr    ee":ws});		
		wb.sheet_names().length.should.equal(3);
	});

	it('Can merge two arraybuffers', function(){
		var a = template.toFile();
		var b = template.toFile();		
		var m = new workbook(mergebooks(a,b));
		m.sheet_names().length.should.equal(2);
	});

	it('Can merge two Workbook objects', function(){
		var a = new workbook(template.toFile());
		var b = new workbook(template.toFile());		
		var m = new workbook(mergebooks(a,b));
		m.sheet_names().length.should.equal(2);
	});
});