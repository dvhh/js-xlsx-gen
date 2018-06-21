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

	let getCellType = function( input ) { return ( ( ( input == "" ) || isNaN( Number( input ) ) ) ? "String" : "Number" ); }
	//let processCellData = function()
	let cell2xml = function( input ) { return ( "<Cell><Data ss:Type=\"" + getCellType( input ) + "\">" + input.encodeHTML() + "</Data></Cell>" ); }

	let row2xml = function( input ) {
		let cells=input.replace("\r",'').split(",").map(cell2xml);
		return( "<Row>"+cells.join("\n")+"</Row>" );
	};
	let csv2table = function( input ) {
		//console.log( input );
		let rows=input.split("\n").map(row2xml);
		return( "<Table><Column ss:Index=\"1\" ss:AutoFitWidth=\"0\" ss:Width=\"110\"/>"+rows.join("\n")+"</Table>" );
	};
	let csv2sheet = function( input ) {
		let table = csv2table( input["content"] );
		return( "<Worksheet ss:Name=\""+input["name"].encodeHTML()+"\">"+table+"</Worksheet>" );
	};
	let csv2xml = function( input ) {
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
		if( navigator.msSaveBlob ) {
			let blob = new Blob( [ unescape(encodeURIComponent(data)) ] ,{"type":"application/vnd.ms-excel" });
			return navigator.msSaveBlob( blob, "workbook.xml");
		}
		let element = document.createElement("a");
		element.download = "workbook.xml";
		element.href="data:application/vnd.ms-excel;base64,"+btoa(unescape(encodeURIComponent(data)));
		element.click();
	}

	{
		let list = document.getElementById( "files" );
		list.addEventListener( "change", function(e) {
			console.log(e);
			let hasSelected = 0;
			for(let i = 0; i < e.target.options.length; ++i ) {
				if( e.target.options[i].selected ) {
					hasSelected = true;
					break;
				}
			}
			document.getElementById( "renameFile" ).disabled = !hasSelected;
			document.getElementById( "deleteFile" ).disabled = !hasSelected;
		} );

		console.log(list);
		// Wire event on "Add file(s)" button
		document.getElementById( "addFile" ).addEventListener( "click", function( e ) {
			let element = document.createElement("input");
			element.type="file";
			element.multiple="multiple";
			element.accept=".csv";
			element.addEventListener( 'change', function( e ) {
				let files = e.target.files;
				for( let i=0; i < files.length; ++i ) {
					let file=files[i];
					let reader = new FileReader( );
					reader.onload = function ( e ) {
						/*contents.push({"name":file.name,"content":e.target.result});
						if( contents.length == files.length ) {
							let result = csv2xml(contents);
							download(result);
						}*/
						console.log(e)
						let option = document.createElement("option");
						option.innerHTML = file.name;
						option.value=e.target.result;
						list.appendChild(option);
						document.getElementById( "convert" ).disabled=false;
					}
					reader.readAsText( file );
				}
			} );
			element.style.display="none";
			document.body.appendChild(element);
			element.click()
		});

		// Wire event on "Convert" button
		document.getElementById( "convert" ).addEventListener( "click", function( e ) {
			let contents = new Array();
			for(let i=0; i < list.options.length; ++i ) {
				contents.push({
					"name"    : list.options[i].outerText,
					"content" : list.options[i].value
				});
			}
			let result = csv2xml(contents);
			download(result);
		});

		// Wire event on "Rename" button
		document.getElementById( "renameFile" ).addEventListener( "click", function( e ) {
			for(let i=0; i < list.options.length; ++i ) {
				if(list.option) {}
			}
		});
	}

	let targets = document.getElementsByClassName("csvConverter");
	//for(let i=0;i < )
	document.getElementsByTagName("div")[0].style.display="block";
})()