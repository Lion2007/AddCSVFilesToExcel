// ************************ Drag and drop ***************** //
let dropArea = document.getElementById("drop-area")

// Prevent default drag behaviors
;['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
  dropArea.addEventListener(eventName, preventDefaults, false)   
  document.body.addEventListener(eventName, preventDefaults, false)
})

// Highlight drop area when item is dragged over it
;['dragenter', 'dragover'].forEach(eventName => {
  dropArea.addEventListener(eventName, highlight, false)
})

;['dragleave', 'drop'].forEach(eventName => {
  dropArea.addEventListener(eventName, unhighlight, false)
})

// Handle dropped files
dropArea.addEventListener('drop', handleDrop, false)

var fileName;
var showout = document.getElementById("showresult");
var re = /(?:\.([^.]+))?$/;
var base_xlsx_file = {};
var result_xlsx_file;
var files_;
var csv_files = [];
var downloadResultCounter;

function preventDefaults (e) {
  e.preventDefault()
  e.stopPropagation()
}

function highlight(e) {
  dropArea.classList.add('highlight')
}

function unhighlight(e) {
  dropArea.classList.remove('active')
}

function handleDrop(e) {
  var dt = e.dataTransfer
  var files = dt.files

  handleFiles(files)
}


function handleFiles(files) {
  files_ = [...files]
  csv_files = [];
  downloadResultCounter = 0;
  //initializeProgress(files.length)
  files_.forEach(do_file)
  //files.forEach(previewFile)
}

const CSVToArray2 = (data, delimiter = ';', omitFirstRow = false) =>
  data
    .slice(omitFirstRow ? data.indexOf('\n') + 1 : 0)
    .split('\n')
    .map(v => v.split(delimiter));

function CSVToArray( strData, strDelimiter ){
	// Check to see if the delimiter is defined. If not,
	// then default to comma.
	strDelimiter = (strDelimiter || ",");

	// Create a regular expression to parse the CSV values.
	var objPattern = new RegExp(
		(
			// Delimiters.
			"(\\" + strDelimiter + "|\\r?\\n|\\r|^)" +

			// Quoted fields.
			"(?:\"([^\"]*(?:\"\"[^\"]*)*)\"|" +

			// Standard fields.
			"([^\"\\" + strDelimiter + "\\r\\n]*))"
		),
		"gi"
		);


	// Create an array to hold our data. Give the array
	// a default empty first row.
	var arrData = [[]];

	// Create an array to hold our individual pattern
	// matching groups.
	var arrMatches = null;


	// Keep looping over the regular expression matches
	// until we can no longer find a match.
	while (arrMatches = objPattern.exec( strData )){

		// Get the delimiter that was found.
		var strMatchedDelimiter = arrMatches[ 1 ];

		// Check to see if the given delimiter has a length
		// (is not the start of string) and if it matches
		// field delimiter. If id does not, then we know
		// that this delimiter is a row delimiter.
		if (
			strMatchedDelimiter.length &&
			strMatchedDelimiter !== strDelimiter
			){

			// Since we have reached a new row of data,
			// add an empty row to our data array.
			arrData.push( [] );

		}

		var strMatchedValue;

		// Now that we have our delimiter out of the way,
		// let's check to see which kind of value we
		// captured (quoted or unquoted).
		if (arrMatches[ 2 ]){

			// We found a quoted value. When we capture
			// this value, unescape any double quotes.
			strMatchedValue = arrMatches[ 2 ].replace(
				new RegExp( "\"\"", "g" ),
				"\""
				);

		} else {

			// We found a non-quoted value.
			strMatchedValue = arrMatches[ 3 ];

		}


		// Now that we have our value string, let's add
		// it to the data array.
		arrData[ arrData.length - 1 ].push( strMatchedValue );
	}

	// Return the parsed data.
	return( arrData );
}


function do_file(file) {
	var f = file;
	var reader = new FileReader();
	var ext = re.exec(file.name)[1];
	
	if (ext.toLowerCase() == "xlsx"){
		reader.onload = function(e) {
			var data = e.target.result;
			data = new Uint8Array(data);
			base_xlsx_file.file = XLSX.read(data, {type:'array'});
			result_xlsx_file = XLSX.read(data, {type:'array'});
			base_xlsx_file.filename = file.name;
			base_xlsx_file.loaded = "yes";
			//process_wb(base_xlsx_file.file);
			process_wb();
		};
		reader.readAsArrayBuffer(f);
	} else if(ext.toLowerCase() == "csv"){
		reader.onload = function(e) {
			var data = e.target.result;
			//alert(file.name);
			csv_files.push({file: data, filename: file.name, loaded: "yes"});
			process_wb();
			
		};
		reader.readAsText(f);
	}
}

function process_wb() {
	// is everything OK then concat
	if(
		(files_ !==null)
		&&(files_.length == csv_files.length +1)
		&&(base_xlsx_file !==null)
		&&(base_xlsx_file.file !==null) 
		&&(base_xlsx_file.loaded == "yes")
		&&(downloadResultCounter == 0)
		) {

		/* get the Gambling worksheet */
		//result_xlsx_file = base_xlsx_file
		var ws = base_xlsx_file.file.Sheets["Gambling"];
          
		  
		try {
		/*  const exceldata =      XLSX.utils.sheet_to_json(ws, {
		   raw: false,
		   header: 1,
		   dateNF: 'yyyy-mm-dd',
		   blankrows: false,
		  });
		*/
		//sorting
		csv_files.sort(function(a, b){
		let x = a.filename;
		let y = b.filename;
		if (x < y) {return -1;}
		if (x > y) {return 1;}
		return 0;
		}); 
		
		csv_files.sort;
		
		//  add data from CSV files
		csv_files.forEach(function(item){
			/*var lines = item.file.split('\n');
			var prevline = "";
			for(var line = 1; line < lines.length; line++){
			linestr = prevline + lines[line];
			
			var columns = linestr.split(',');
			
			// check lines without data
			if(columns.length == 1){
				continue;
			}
			
			var array = CSVToArray(linestr, ";");
			exceldata.push({Service: array[0][3]});
			};*/
			array = CSVToArray2(item.file);
			
			// check last line
			var lines = item.file.split('\n');
			if(lines[lines.length-1] == ''){
				array.splice(lines.length-1, 1)
			}
			//del headers
			array.splice(0, 1)
			
			data_ = item.filename.replace( /^\D+/g, '');
			
			
			for (var i = 0; i < array.length; i++) {
				// del quotes
				array[i][0] = array[i][0].replaceAll('"', '');
				// replace . to ,
				array[i][1]  =  array[i][1].replaceAll('.', ',');
				array[i][2]  =  array[i][2].replaceAll('.', ',');
				array[i][3]  =  array[i][3].replaceAll('.', ',');
				array[i][4]  =  array[i][4].replaceAll('.', ',');
				array[i][5]  =  array[i][5].replaceAll('.', ',');
				array[i][6]  =  array[i][6].replaceAll('.', ',');
				array[i][7]  =  array[i][7].replaceAll('.', ',');
				array[i][8]  =  array[i][8].replaceAll('.', ',');
				array[i][9]  =  array[i][9].replaceAll('.', ',');
				array[i][10] = array[i][10].replaceAll('.', ',');
				// del %
				array[i][10] = array[i][10].replaceAll('%', '');
				// add_date
				array[i][11] = data_.substr(8, 2);
				array[i][12] = data_.substr(5, 2);
				array[i][13] = data_.substr(0, 4);
				array[i][14] = '';
			}
			//XLSX.WritingOptions.cellN = true;
			XLSX.SSF.format('$#,##0.00', 12345.6789)
			XLSX.utils.sheet_add_aoa(ws, array, { origin: -1 , cellNF: true, cellText: false, cellStyles: true, raw:false});
		
		});
		
		var colNum = XLSX.utils.decode_col("B"); //decode_col converts Excel col name to an integer for col #
		var fmt = '$0.00'; // or '"$"#,##0.00_);[Red]\\("$"#,##0.00\\)' or any Excel number format

		/* get worksheet range */
		var range = XLSX.utils.decode_range(ws['!ref']);
		for(var i = range.s.r + 1; i <= range.e.r; ++i) {
		  /* find the data cell (range.s.r + 1 skips the header row of the worksheet) */
		  var ref = XLSX.utils.encode_cell({r:i, c:colNum});
		  /* if the particular row did not contain data for the column, the cell will not be generated */
		  if(!ws[ref]) continue;
		  /* `.t == "n"` for number cells */
		  if(ws[ref].t != 'n') continue;
		  /* assign the `.z` number format */
		  ws[ref].z = fmt;
		}
		
		
		//const worksheet = XLSX.utils.json_to_sheet(exceldata);

		//const workbook = XLSX.utils.book_new();
		//XLSX.utils.book_append_sheet(base_xlsx_file.file, worksheet, "Gambl1ing_new!");
		
		  
		
		
		//Service	Unique players	Sessions	Rounds	Betting slips	Bet	Win	GGR	JP Cont,	JP win	Margin	Day	Month	Year	Brand
		//Service	Unique players	Sessions	Rounds	Betting slips	Bet	Win	GGR	JP Cont.	JP win	Margin


		  
			//alert(exceldata);
		 } catch (e) {
		  showout.innerHTML  = 'Error ' + e.name + ":" + e.message + "    " + e.stack;
		 }
		/* save file */
		XLSX.writeFile(base_xlsx_file.file, "issue1124.xlsx");
	}
}

function readFile(file) {
	
  const reader = new FileReader();
  const ext = re.exec(file.name)[1];
  
  reader.addEventListener('load', (event) => {
	try{
		var result = event.target.result;
		fileName = file.name;
		
		if (ext.toLowerCase() == "xls"){
			base_xlsx_file = XLSX.read(result, {type: 'array'});
			//result = XLS.utils.make_csv(base_xlsx_file.Sheets[base_xlsx_file.SheetNames[0]]); 
			// Process Data (add a new row)
			//var ws = base_xlsx_file.Sheets["Gambling"];
			//XLSX.utils.sheet_add_aoa(ws, [["Created "+new Date().toISOString()]], {origin:-1});

			XLSX.writeFile(base_xlsx_file, "Report.xlsx");


			
		}
		
		/*var string = "";
		if(result.split('\n')[0] == '"Posting Date","Value Date","UTN","Description","Debit","Credit","Balance"'){
			string = convertEurobank(result);//"Posting Date,Value Date,UTN,Description,Payee,Debit_Credit,Balance\n";
		}else if(result.split('\n')[0] == 'reference,datetime,valuedate,debit,credit,trname,contragent,rem_i,rem_ii,rem_iii\r'){
			string = convertFIBank(result);
		}else if(result.split('\n')[0] == '"Account owner","Account number","Account type",Currency,Description,Balance'){
			string = convertPaymentExecution(result);
		}else if(result.split('\n')[0] == ',,,,,THE LUCK FACTORY EUROPE LTD,,,'){
			string = convertEcommBX(result);
		}else if((ext == "xls")&&(result.split('\n')[0] == 'ACCOUNT NO,PERIOD,CURRENCY,DATE,DESCRIPTION,DEBIT,CREDIT,VALUE DATE,BALANCE')){
			string = convertHellenic(result);//"Posting Date,Value Date,UTN,Description,Payee,Debit_Credit,Balance\n";
		}else{
			showout.innerHTML  = "Do not recognize the bank";
			return;
		}*/
		
		
		//doSave(string, fileName, ext);
	} catch(e) {
			showout.innerHTML  = 'Error ' + e.name + ":" + e.message + "    " + e.stack;
	}
  });
  
  reader.addEventListener('error', (event) => {
	  showout.innerHTML  = reader.error; 
    });
  if ((ext.toLowerCase() == "xls")||(ext.toLowerCase() == "xlsx")){
	reader.readAsBinaryString(file);  
  }	else{
	reader.readAsText(file);  
  }
  
}

function doSave(content, filename, ext) {
	var today = new Date();
	var date = today.getFullYear() + String((today.getMonth()+1)).padStart(2, '0') + String(today.getDate()).padStart(2, '0');
	var time = String(today.getHours()).padStart(2, '0')  + String(today.getMinutes()).padStart(2, '0')  + String(today.getSeconds()).padStart(2, '0');
	var dateTime = date+time;
	
	fileName = filename.slice(0,filename.length - ext.length -1) + '_converted_at_' + dateTime + '.csv';
	
    var blob = new Blob([content], {
        type: "data:text/plain;charset=utf-8"
    });
	
    saveAs(blob, fileName);
	date = today.getFullYear()+'.'+String((today.getMonth()+1)).padStart(2, '0')+'.'+String(today.getDate()).padStart(2, '0');
	time = today.getHours() + ":" + String(today.getMinutes()).padStart(2, '0') + ":" + String(today.getSeconds()).padStart(2, '0');
	dateTime = date+" "+time;
	showout.innerHTML  = fileName + " successfully converted at " + dateTime; 
}



