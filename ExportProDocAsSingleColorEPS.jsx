// author: Miquel Brazil
// version: 0.1.0

var documentColors = [];
var refNumber = null;
var shipDate = null;
var debug = false;
var date = new Date();

if (debug) {
	refNumber = '18-12345';
	shipDate = '09/15';
} else {
	refNumber = prompt("Enter the Job Number for this ProDoc", '', "ProDoc Job Number");
	shipDate = prompt("Enter Ship Date for this Job in MM/DD format", '', "ProDoc Ship Date");
}

if (refNumber && shipDate) {
	$.writeln("The Job Number is " + refNumber);
	$.writeln("The Ship Date is " + shipDate);

	var doc = activeDocument;

	try {
		var artworkLayer = doc.layers.getByName('artwork');
		$.writeln("Artwork layer was found.");

		var templateLayer = doc.layers.getByName('template');


		templateLayer.visible = false;
		$.writeln("Turned off the Template layer");

	} catch (e) {
		$.writeln(e);
		alert("No Artwork layer was detected.");
	}

	try {
		inspect(artworkLayer);
	} catch (e) {
		alert(e.toString());
	}

	/*for (var s = 0; s < doc.spots.length; s++) {
		$.writeln(doc.spots[s].name);
	}*/

	var colors = "";


	doc.characterStyles.removeAll();

	for(var c = 0; c < documentColors.length; c++) {
		colors = colors + documentColors[c] + " ";

		$.writeln("===========");
		$.writeln(documentColors[c]);
		var style = doc.characterStyles.add(documentColors[c]);
		var styleAttr = style.characterAttributes;

		if (documentColors[c] !== "Black") {
			var spot = new SpotColor();
			spot.spot = doc.spots.getByName(documentColors[c]);
		} else {
			var spot = new GrayColor();
			spot.gray = 100;
		}

		var none = new NoColor();

		styleAttr.capitalization = FontCapsOption.ALLCAPS;
		styleAttr.fillColor = spot;
		styleAttr.strokeColor = none;
	}

	$.writeln(colors.length);

	var infoLayer = doc.layers.getByName('info');

	infoLayer.groupItems.getByName('Job').textFrames.getByName('meta').lines[1].contents = colors;

	//documentColors = documentColors.reverse();
	var startChar = 0;
	var endChar = 0;
	for(var c = 0; c < documentColors.length; c++) {
		if (c > 0) {
			startChar = documentColors[c - 1].length + startChar + 1;
		}

		endChar = startChar + documentColors[c].length + 1;

		$.writeln("The style application would start on character " + startChar + " and end on character " + endChar);

		for(var t = startChar; t < endChar; t++) {
			$.writeln(infoLayer.groupItems.getByName('Job').textFrames.getByName('meta').lines[1].characters[t].contents);
			doc.characterStyles.getByName(documentColors[c]).applyTo(infoLayer.groupItems.getByName('Job').textFrames.getByName('meta').lines[1].characters[t], true);
		}

		$.writeln("======================")
	}


	$.writeln(infoLayer.groupItems.getByName('Job').textFrames.getByName('meta').lines[0].contents);
	infoLayer.groupItems.getByName('Job').textFrames.getByName('meta').lines[0].contents = refNumber;


	function inspect(thePageItems) {
		for(var i = 0; i < thePageItems.pageItems.length; i++) {
    		var item = thePageItems.pageItems[i];
    		if (item.typename === "GroupItem" && !item.guides) {
				inspect(item);
			} else if (item.typename === "PlacedItem" || item.typename === "RasterItem") {
				$.writeln(item.typename);
				throw new Error("This ProDoc contains Raster images and cannot be saved. Please remove the Raster images or convert them to vector and try again.");
			} else if ((item.fillColor || item.strokeColor) && !item.guides) {
				getColors(item);
			} else if (item.typename === "TextFrame") {
				getColors(item.textRange.characterAttributes);
			} else if (item.typename === "CompoundPathItem") {
				for (var c = 0; c < item.pathItems.length; c++) {
					getColors(item.pathItems[c]);
				}
			} else {
    			$.writeln(item.typename);
    		}
		}

		return;
	}

	$.writeln(documentColors.toString());

	/*

	var datePath = date.getFullYear().toString() + ('0' + (date.getMonth()+1).toString()).slice(-2) + ('0' + date.getDate().toString()).slice(-2);
	var loc = new Folder(Folder.desktop + "/Leprechaun/" + datePath);

	if (!loc.exists) {
		loc.create();
	}

	$.writeln(loc.exists);

	$.writeln(documentColors.toString());

	// save document

	var file = new File(loc.fsName + "/" + refNumber);
	$.writeln(file);

	var epsOptions = new EPSSaveOptions();
	epsOptions.embedAllFonts = true;
	epsOptions.artboardRange = '1';
	epsOptions.saveMultipleArtboards = true;

	doc.saveAs(file, epsOptions);

	*/

} else {
	alert("You must provide a Job Number and Ship Date to continue.");
}


function getColors(item) {
	try {
		var fillColorType = item.fillColor.typename;
		var strokeColorType = item.strokeColor.typename;
		$.writeln("Fill Color Type for " + item.typename + ' is ' + fillColorType);

		switch (fillColorType) {
			case "CMYKColor":
				if (item.fillColor.cyan === 0 && item.fillColor.magenta === 0 && item.fillColor.yellow === 0) {
					if (item.fillColor.black > 0) {
						if (documentColors.toString().indexOf("Black") === -1) {
							documentColors.push("Black");
						}
					}
				} else {
					throw new Error("ProDoc contains invalid CMYK Process Colors. Please check Artwork and try again.")
				}
				break;
			case "GrayColor":
				if (item.fillColor.gray > 0) {
					if (documentColors.toString().indexOf("Black") === -1) {
						documentColors.push("Black");
					}
				}
				break;
			case "SpotColor":
				if (documentColors.toString().indexOf(item.fillColor.spot.name) === -1) {
					documentColors.push(item.fillColor.spot.name);
				}
				break;
			case "NoColor":
				break;
		}
	} catch (e) {
		alert(e.toString())
		return;
	}
}
