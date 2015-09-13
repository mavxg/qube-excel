"use strict";
var XLSX = require('xlsx');
var JSZip = require('jszip');

function Workbook(content){
	if(!(this instanceof Workbook)) return new Workbook(content);
	if(content){
		if(content instanceof ArrayBuffer)	{
			var bstr = utils.arrayBufferToBinaryString(content);
			this.zip = new JSZip(bstr, { base64:false });	
		} else { //Must be an object of sheets
			this.Sheets = {};
			this.SheetNames = [];
			this.write(content);		
		}
	}	
}

Workbook.prototype.write = function(sheets){
	var wopts = { bookType:'xlsx', bookSST:false, type:'binary' };
	for(var prop in sheets){
		if(sheets.hasOwnProperty(prop)){
			this.SheetNames.push(prop);
			this.Sheets[prop] = sheets[prop];
		}
	}		
	this.zip = new JSZip(XLSX.write(this, wopts), { base64:false });
};

Workbook.prototype.read_in_workbook = function(content){
	var bStr = utils.arrayBufferToBinaryString(content);
	this.zip = new JSZip(bStr, { base64:false });
}

Workbook.prototype.read_master_sheet = function(){
	var sheet = XLSX.readSheet(this.zip, 'xl/Workbook.xml');
	return sheet.match(/sheet name=\"([\w\s]*)\"\ssheetId=\"[0-9]*\"\sr:id=\"rId[0-9]*\"/gi).map(function(d){
		return { name:d.match(/name=\"([\w\s]*)\"/)[1], 
		         id:d.match(/rId(\w*)\"/)[1]}
	});
};

Workbook.prototype.read_sheet = function(sheet_path){
	return XLSX.readSheet(this.zip, sheet_path);
};

Workbook.prototype.sheet_names = function(){
	var sheets = this.read_master_sheet();
	return sheets.map(function(sheet){
		return sheet.name;
	});
}

Workbook.prototype.sheet_aliases = function(){
	var content = this.read_sheet('[Content_Types].xml')
	return content.match(/sheet[0-9]+.xml/g);
}

Workbook.prototype.table_aliases = function(){
	var content = this.read_sheet('[Content_Types].xml')
	return content.match(/table[0-9]+.xml/g);
}

function next_free(collectionWithNumberInStr){	
	var mx = 1;
	if(!collectionWithNumberInStr) return mx;
	collectionWithNumberInStr.forEach(function(s){
		var num = parseInt(s.match(/[0-9]+/)[0]);
		mx = (num>mx) ? num : mx;
	});
	return mx + 1;
}

Workbook.prototype.next_free_sheet = function(){
	return 'xl/worksheets/sheet' + next_free(this.sheet_aliases()) + '.xml';
}

Workbook.prototype.next_free_table = function(){
	return 'xl/tables/table'+next_free(this.table_aliases()) + '.xml';
}

Workbook.prototype.find_sheet_path = function(name){
	var sheets = this.read_master_sheet();
	for (var i = sheets.length - 1; i >= 0; i--)
		if(name.toLowerCase() == sheets[i].name.toLowerCase())
			return 'xl/worksheets/sheet' + sheets[i].id + '.xml';
	return undefined;
};

Workbook.prototype.find_table_path = function(sheetName){
	var alias = this.find_sheet_path(sheetName);
	var sheetAlias = alias.match(/([A-z0-9]*)\.xml/)[1];
	var dataTableName = 'xl/worksheets/_rels/' + sheetAlias + '.xml.rels';
	if(this.zip.files.hasOwnProperty(dataTableName)) {
		var relsContent = this.read_sheet(dataTableName);
		return 'xl/tables/' + relsContent.match(/[a-z0-9]*\.xml/g)[0];
	}
	return undefined;
}

Workbook.prototype.sheet_datarange = function(sheetData){
	return /\<dimension ref="([\w\d:]*)"\/\>/gm.exec(sheetData)[1];
};

//Reassemble the zip
Workbook.prototype.toFile = function(){
	return utils.stringToArrayBuffer(this.zip.generate({type:"string"}));
};

Workbook.prototype.clean_cache = function(sheetNames){
	//If cell has a <c> then <f> then <v>, remove the value in <v>, should cause Excel to recalc on open
	
	var wb = this;
	/**
	sheetNames = sheetNames || wb.sheet_names();
	sheetNames.forEach(function(sheet){
		var alias = wb.find_sheet_path(sheet);
		var data = wb.read_sheet(alias);
		try {
		var trimmed = data.replace(/(\<c\s.+?\>\<f.+?[^]+?\<v\>)(.+?)(\<\/v\>)/g, "$1$3");
		if(data != trimmed)
			wb.zip.files[alias]._data = utils.stringToUint8Array(trimmed);
		} catch(err) {
			console.log(err.toString());
		}	
	});**/

	//remove the calcChain as formulae will probably be inconsistent.  Excel recreates on open anyway
	if(wb.zip.files.hasOwnProperty('xl/calcChain.xml'))
		delete wb.zip.files['xl/calcChain.xml'];
}

Workbook.prototype.replace_sheet_data = function(sheetName, data, range){
	var alias = this.find_sheet_path(sheetName);
	var sheet = this.read_sheet(alias);

	sheet = sheet.replace(/\<sheetData\/\>/g, data); //Replace an empty data tag
	sheet = sheet.replace(/\<sheetData.*sheetData\>/g, data); //Replace the data
	sheet = sheet.replace(/\<dimension ref="[\w\d:]*"\/\>/g, "<dimension ref=\""+range+"\"/>"); //Replace the range

	this.zip.files[alias]._data = utils.stringToUint8Array(sheet);
}

Workbook.prototype.replace_sheet_contents = function(sheetName, dataSheet){	
	this.zip.files[sheetName]._data = utils.stringToUint8Array(dataSheet);
}

Workbook.prototype.has_datatable = function(sheetName){
	var alias = this.find_sheet_path(sheetName);
	var sheetAlias = alias.match(/([A-z0-9]*)\.xml/)[1];
	var dataTableName = 'xl/worksheets/_rels/' + sheetAlias + '.xml.rels';
	return this.zip.files.hasOwnProperty(dataTableName);
}

//Produce a new worksheet
//headers = []
//rows = [[]]
//keyCount = int
//params = {header:bool, startCell:"A1", skipEmptyRows:bool, colWidth:int, headers:["",""]}
function Worksheet(headers, rows, keyCount, params){
	function addCell(elem, row, col){
		if(!elem) elem = '';

		var cell = {v:elem};

		if(typeof cell.v === 'number') {
			if(isNaN(cell.v)){
				cell.v = '';
				cell.t = 's';
			}
			else {
				cell.t = 'n';
				cell.w = elem.toLocaleString();
			}
		}
		else if(typeof cell.v === 'boolean') cell.t = 'b';
		else if(cell.v instanceof Date) {
			cell.t = 'n'; 
			cell.z = XLSX.SSF._table[14];
			cell.v = datenum(cell.v);
		}
		else cell.t = 's';
		cell_ref = XLSX.utils.encode_cell({c:col,r:row});
		ws[cell_ref] = cell;		
	}	
	params = params || {};
	var header    = utils.parseBoolean(params.header, true);
	var skipEmpty = utils.parseBoolean(params.skipEmptyRows, false);
	var cell_ref  = params.startCell || 'A1';
	var colWidth  = parseInt(params.colWidth || 20);
	var rOffset   = header ? 1 : 0;	
	var start_pos = XLSX.utils.decode_cell(cell_ref);
	var ws        = {};
	var colWidths = [];

	var rowNum = start_pos.r;	

	headers = (params.headers || headers);
	headers.forEach(function(val,col){
		if(header) addCell(val,rowNum,col+start_pos.c);
		colWidths.push({wch:colWidth});
	});
	
	rows.forEach(function(row){
		//If you want to skip emptys and every non key elem is empty
		if(skipEmpty && row.slice(keyCount).every(function(elem){ return (elem ? false : true) })){
			//Skip
		} else {
			if(Array.isArray(row)){
				row.forEach(function(elem, colNum){
				addCell(elem,rowNum+rOffset,colNum+start_pos.c);
				});	
			} else {
				addCell(row,rowNum+rOffset,start_pos.c);
			}
			rowNum++;
		}
	});

	ws['!ref'] = 'A1:'+cell_ref;
	ws['!cols'] = colWidths;
	return ws;
}

var utils = {
	stringToArrayBuffer: function(s){
		var ab = new ArrayBuffer(s.length);
		var view = new Uint8Array(ab);
		for (var i=0; i!=s.length; ++i) 
			view[i] = s.charCodeAt(i) & 0xFF;
		return ab;
	},
	stringToUint8Array: function(s){
		var uint = new Uint8Array(s.length);
		for(var i=0,j=s.length;i<j;++i)
			uint[i]=s.charCodeAt(i);
		return uint;
	},
	arrayBufferToBinaryString: function(ab) {
		var data = new Uint8Array(ab);
		var arr = new Array();
		for(var i=0; i!=data.length; ++i)
			arr[i] = String.fromCharCode(data[i]);
		return arr.join("");
	},
	parseBoolean: function(val, fallback){
		if(typeof val === 'boolean') return val;
		if(typeof val === 'number') {
			if(val % 1 == 0 && val <= 1 && val >= 0) return Boolean(val);
		}
		if(val!=undefined){
			var str = val.toLowerCase();
			if(str == "true" || str == "1") return true;
			if(str == "false" || str == "0") return false;
		}
		return fallback;
	}
}

module.exports = {
	Workbook:Workbook,
	Worksheet:Worksheet,
};