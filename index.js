(function(){
	"use strict";

	/*------------------------------------------------------------------------*/

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

	/*------------------------------------------------------------------------*/

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
	/*------------------------------------------------------------------------*/
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

	/*------------------------------------------------------------------------*/
	let setupButtonEventListener = function( targetId, event, handler ) {
		if( typeof(targetId) == "undefined" ) { return; }
		let button = document.getElementById(targetId);
		if( typeof(button) == "undefined" ) { return; }
		button.addEventListener( event, handler );
	};

	let disableButton = function ( targetId, disabled ) {
		if( typeof(targetId) == "undefined" ) { return; }
		let button = document.getElementById(targetId);
		if( typeof(button) == "undefined" ) { return; }
		button.disabled = disabled;
	};

	let setupList = function (target) {
		console.log(target.dataset);

		const buttonConvertId    = target.dataset["buttonConvertId"];
		const buttonDeletefileId = target.dataset["buttonDeletefileId"];
		const buttonRenamefileId = target.dataset["buttonRenamefileId"];

		target.addEventListener( "change", function(e){
			let hasSelected = 0;
			for(let i = 0; i < e.target.options.length; ++i ) {
				if( e.target.options[i].selected ) {
					hasSelected = true;
					break;
				}
			}

			[
				buttonDeletefileId,
				buttonRenamefileId
			].forEach(
				function(key) {
					disableButton( key, !hasSelected )
				}
			);
		});

		setupButtonEventListener(
			target.dataset["buttonAddfileId"],
			"click",
			function ( e ) {
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
							console.log(e)
							let option = document.createElement("option");
							option.innerHTML = file.name;
							option.value=e.target.result;
							target.appendChild(option);
							disableButton( buttonConvertId , false );
						}
						reader.readAsText( file );
					}
				} );
				element.style.display="none";
				document.body.appendChild(element);
				element.click();
			}
		);

		setupButtonEventListener(
			buttonConvertId,
			"click",
			function( e ) {
				let contents = new Array();
				for(let i=0; i < target.options.length; ++i ) {
					contents.push({
						"name"    : target.options[i].outerText,
						"content" : target.options[i].value
					});
				}
				let result = csv2xml(contents);
				download(result);
			}
		);

		setupButtonEventListener(
			buttonDeletefileId,
			"click",
			function( e ) {
				let toRemove = new Array();
				for(let i = 0; i < target.options.length; ++i ) {
					if( target.options[i].selected && window.confirm("delete '"+ target.options[i].outerText + "' ?" ) ) {
						toRemove.push(target.options[i]);
					}
				}
				toRemove.forEach( function(element) { target.removeChild(element); });
			}
		);

		setupButtonEventListener(
			buttonRenamefileId,
			"click",
			function( e ) {
				for(let i = 0; i < target.options.length; ++i ) {
					if( target.options[i].selected  ) {
						target.options[i].innerHTML = window.prompt("rename '" + target.options[i].outerText + "' to ?", target.options[i].outerText );
					}
				}
			}
		);
	};

	/*------------------------------------------------------------------------*/
	let targets = document.getElementsByClassName("csvConverter");
	console.log(targets);

	for( let i=0;i < targets.length; ++i ) {
		setupList(targets[i])
	}

	/*------------------------------------------------------------------------*/
	document.getElementsByTagName("div")[0].style.display = "block";
})()