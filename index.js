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

			//Replace the data
			tb.replace_sheet_contents(sheetName, dataSheet);

			//If there is a datatable referencing the sheet replace the range, remove any sorts
			if(tb.has_datatable(sheetName)){

				//Think this is wrong as we need to get the table.xml
				var templateXml = tb.read_sheet(templatePath);
				templateXml = templateXml.replace(/ref\=\"[A-z0-9:]*\"/gm,"ref=\""+range+"\"");
				templateXml = templateXml.replace(/<sort.*State>/,"");

				//Not sheetname...?
				tb.replace_sheet_contents(sheetName, templateXml);
			}
		} 
		//Clear the template cache
		tb.clean_cache();			
	});
	return tb.toFile();
}

module.exports = {
	Workbook: Workbook,	
	Worksheet: excel.Worksheet,	
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