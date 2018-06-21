(function(){
	"use strict";
	if (!String.prototype.encodeHTML) {
		String.prototype.encodeHTML = function () {
			return this
				.replace(/&/g, '&amp;')
				.replace(/</g, '&lt;')
				.replace(/>/g, '&gt;')
				.replace(/"/g, '&quot;')
				.replace(/'/g, '&apos;');
		};
	}

	let cell2xml = input => "<Cell><Data ss:Type=\"String\">"+input.encodeHTML()+"</Data></Cell>";

	let row2xml = function( input ) {
		let cells=input.replace("\r",'').split(",").map(cell2xml);
		return("<Row>"+cells.join("\n")+"</Row>");
	}
	let csv2table = function( input ) {
		//console.log( input );
		let rows=input.split("\n").map(row2xml);
		return("<Table><Column ss:Index=\"1\" ss:AutoFitWidth=\"0\" ss:Width=\"110\"/>"+rows.join("\n")+"</Table>");
	};
	let csv2sheet = function( input ) {
		let table = csv2table( input["content"] );
		return "<Worksheet ss:Name=\""+input["name"].encodeHTML()+"\">"+table+"</Worksheet>";
	};
	let csv2xml = function( input ) {
		//console.log(input);
		//console.log(document.getElementById('workbook').innerHTML);
		//let result=csv2tables;
		let sheets = input.map(csv2sheet);
		console.log(sheets);
		return( 
				"<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n"+
				"<?mso-application progid=\"Excel.Sheet\"?>\n"+
				"<Workbook xmlns=\"urn:schemas-microsoft-com:office:spreadsheet\" xmlns:x=\"urn:schemas-microsoft-com:office:excel\" xmlns:ss=\"urn:schemas-microsoft-com:office:spreadsheet\" xmlns:html=\"http://www.w3.org/TR/REC-html40\">\n"+
				sheets.join("\n")+
				"</Workbook>"
		);
	};

	let download = function( data ) {
		console.log(data);
		let element = document.createElement("a");
		element.download = "result.xml";
		element.href="data:application/vnd.ms-excel;base64,"+btoa(unescape(encodeURIComponent(data)));
		element.click();
	}

	let fileInput = document.getElementsByTagName("input")[0];
	fileInput.addEventListener( 'change', function( e ) {
		let contents = new Array();

		let files = e.target.files;
		for( let i=0; i < files.length; ++i ) {
			let file=files[i];
			let reader = new FileReader( );
			reader.onload = function ( e ) {
				contents.push({"name":file.name,"content":e.target.result});
				if( contents.length == files.length ) {
					let result = csv2xml(contents);
					download(result);
				}
			}
			reader.readAsText( file );
		}

	} );
})()