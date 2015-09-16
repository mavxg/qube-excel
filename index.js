"use strict";
var excel = require('./lib/js-xlsx-templater');

function Workbook(data){	
	return MergeBooks(data);
}

//Expects data as an object with property names as sheet names and values as Worksheets
//template is another workbook(array buffer)
function MergeBooks(data, template) {
	var wb = (data instanceof excel.Workbook) ? data : new excel.Workbook(data);
	if(!template) return wb.toFile();
	
	var tb = (template instanceof excel.Workbook) ? template : new excel.Workbook(template);
	
	wb.sheet_names().forEach(function(sheetName){
		//If you find the sheetname in the template, switch out the data
		var templatePath = tb.find_sheet_path(sheetName);
		if(templatePath){
			var dataSheet = wb.read_sheet(wb.find_sheet_path(sheetName));
			var range = wb.sheet_datarange(dataSheet);
			var sheetData = dataSheet.match(/\<sheetData.*sheetData\>/g)[0];
			tb.replace_sheet_data(sheetName,sheetData,range);
			
			//If there is a datatable referencing the sheet replace the range, remove any sorts
			var tablePath = tb.find_table_path(sheetName);
			if(tablePath) {			
				var templateXml = tb.read_sheet(tablePath);
				templateXml = templateXml.replace(/ref\=\"[A-z0-9:]*\"/g,"ref=\""+range+"\"");
				templateXml = templateXml.replace(/<sort.*State>/,"");				
				tb.replace_sheet_contents(tablePath, templateXml);
			}
		}					
	});
	//Clear the template cache
	//tb.clean_cache();
	//Set calc properties
	tb.set_calc_properties('fullCalcOnLoad="1"');
	return tb.toFile();
}

function Worksheet(data, params){
	var headers = data.headers;
	var dimensions = (data.dimensions || []).length;
	return excel.Worksheet(headers, data, dimensions, params);
}

module.exports = {
	Workbook: Workbook,	
	Worksheet: Worksheet,	
	MergeBooks: MergeBooks,	
};

/**
TODO: Add new sheets to the template
//Create new reference in workbook.xml.rels
//[Content_Types].xml needs to have entries for sheet{0-9}.xml
//<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
//[Content_Types].xml needs to have entries for tables{0-9}.xml
//<Override PartName="/xl/tables/table1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>
//If adding table need to have _rels to marry the two
**/