/**************************************************************************

Export ProDoc to EPS
ExportProDocEPS.jsx

author: Miquel Brazil
version: 1.0.beta

DESCRIPTION

This script automates the export of ProDoc files after
Artwork has been setup.

It will do the following tasks:

* detects the Colors used by elements on the Artwork layer
* exits an import if document contains CMYK colors or Raster Artwork
* hides the template layer
* write the Job Number to the MetaData field
* write all Colors identified to the MetaData field and sets swatches
* exports to EPS file format against the Artboard appropriate for the
total Imprint Colors detected
* positions the MetaData field according to the Artboard selected
* formats the filename according the the RefNumber_Version_Artboard format

**************************************************************************/

// allows bypass of some UI interactions while developing
var debug = false;

// setup global variables to use across script
var doc = activeDocument;
var documentColors = [];
var strDocumentColors = "";
var refNumber = getReferenceNumber();

// setup Date variables
var date = new Date();
var dateYear = date.getFullYear().toString();
var dateMonth = ('0' + (date.getMonth()+1).toString()).slice(-2);
var dateDay = ('0' + date.getDate().toString()).slice(-2);

var datePath = dateYear + dateMonth + dateDay;
var prodocVersion = getUnixTimestamp(date.getTime());

// setup Fonts
var displayFont = "ArialMT";
var defaultFont = "MyriadPro-Regular";

var noColor = new NoColor();
var registrationColor = new SpotColor();
var registrationColorName = ["[Registration]", "Fire Red C"];
registrationColor.spot = activeDocument.spots.getByName(registrationColorName[0]);

if (refNumber) {
	try {
		var artworkLayer = doc.layers.getByName('artwork');
		var templateLayer = doc.layers.getByName('template');
		templateLayer.visible = false;
	} catch (e) {
		alert("No Artwork layer was detected.");
	}

    if (artworkLayer) {
        try {
            setupProDoc();

            // save ProDoc
        	if (exportEPS()) {
                alert("ProDoc Export is complete.");
            } else {
                alert("There was an issue saving the ProDoc file.");
            }
    	} catch (e) {
    		alert(e.toString());
    	}
    }
} else {
	alert("You must provide a valid Job Number to use this tool.");
}

function setupProDoc() {
    inspect(artworkLayer);

    // sort the identified Colors in alphabetical order
    documentColors.sort();

    // remove any existing Character Styles
    doc.characterStyles.removeAll();

    // setup new Character Styles based on identified Colors
    for (var c = 0; c < documentColors.length; c++) {
        setCharacterStyle(documentColors[c]);
    }

    writeDocumentColors();
    writeArtboardImprintTotal();
    writeReferenceNumber();
    setMetaDataPosition();
}

function inspect(thePageItems) {
    for(var i = 0; i < thePageItems.pageItems.length; i++) {
        var item = thePageItems.pageItems[i];
        if (item.typename === "GroupItem" && !item.guides) {
            inspect(item);
        } else if (item.typename === "PlacedItem" || item.typename === "RasterItem") {
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
            // reserved section for adding future detection types
        }
    }

    return;
}

function getReferenceNumber() {
    var refNumber = null;

    if (debug) {
    	refNumber = '18-23456';
    } else {
        for (var attempt = 1; attempt <= 3; attempt++) {
            var promptMsg = '';
            if (attempt == 1) {
                promptMsg =  "Enter the Job Number for this ProDoc"
            } else {
                promptMsg =  "You entered an invalid Job Number for this ProDoc. Please enter a valid Job Number and try again."
            }

            response = prompt(promptMsg, '', "ProDoc Job Number");

            if (validateReferenceNumber(response)) {
                refNumber = response;
                break;
            }
        }
    }

    return refNumber;
}

function writeReferenceNumber() {
    var refNumberText = doc.layers.getByName('info').groupItems.getByName('job').textFrames.getByName('meta').lines[0];

	refNumberText.contents = refNumber;
    refNumberText.textRanges[0].lines[0].characterAttributes.fillColor = registrationColor;
}

function validateReferenceNumber(refNumber) {
    var validReferenceNumber = false;

    if (refNumber) {
        var regex = /1[56789]-\d{1,5}/;
        validReferenceNumber = regex.test(refNumber);
    }

    return validReferenceNumber;
}

function getColors(item) {
	try {
        // items can have Fill or Stroke Colors
		var fillColorType = item.fillColor.typename;
		var strokeColorType = item.strokeColor.typename;

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

        switch (strokeColorType) {
			case "CMYKColor":
				if (item.strokeColor.cyan === 0 && item.strokeColor.magenta === 0 && item.strokeColor.yellow === 0) {
					if (item.strokeColor.black > 0) {
						if (documentColors.toString().indexOf("Black") === -1) {
							documentColors.push("Black");
						}
					}
				} else {
					throw new Error("ProDoc contains invalid CMYK Process Colors. Please check Artwork and try again.")
				}
				break;
			case "GrayColor":
				if (item.strokeColor.gray > 0) {
					if (documentColors.toString().indexOf("Black") === -1) {
						documentColors.push("Black");
					}
				}
				break;
			case "SpotColor":
				if (documentColors.toString().indexOf(item.strokeColor.spot.name) === -1) {
					documentColors.push(item.strokeColor.spot.name);
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

function getUnixTimestamp(ts) {
    return (ts/1000|0).toString();
}

function exportEPS() {
    // set the Artboard to save against based on the total Imprint Colors detected
    if (documentColors.length == 1) {
        var artboard = '1';
    } else {
        var artboard = '2';
    }

    // build the folder location to save the ProDoc
    var loc = new Folder(Folder.desktop + "/Leprechaun/" + datePath);

    // if the folder hierarchy doesn't exist then create it
	if (!loc.exists) {
		loc.create();
	}

	// save document
	var file = new File(loc.fsName + "/" + refNumber + "_" + prodocVersion);

	var epsOptions = new EPSSaveOptions();
	epsOptions.embedAllFonts = true;
    epsOptions.saveMultipleArtboards = true;
	epsOptions.artboardRange = artboard;

	doc.saveAs(file, epsOptions);

    return true;
}

function setCharacterStyle(color) {
    // setup new Character Styles based on identified Artwork colors
    // build a string with all identified Artwork colors
    strDocumentColors = strDocumentColors + color + " ";

    var style = doc.characterStyles.add(color);
    var styleAttr = style.characterAttributes;

    if (color !== "Black") {
        var spotColor = new SpotColor();
        spotColor.spot = doc.spots.getByName(color);
    } else {
        var spotColor = new GrayColor();
        spotColor.gray = 100;
    }

    try {
        var font = textFonts.getByName(displayFont);
    } catch (e) {
        var font = textFonts.getByName(defaultFont);
    }

    styleAttr.textFont = font;
    styleAttr.capitalization = FontCapsOption.ALLCAPS;
    styleAttr.fillColor = spotColor;
    styleAttr.strokeColor = noColor;
}

function writeArtboardImprintTotal() {
    if (documentColors.length == 1) {
        var artboard = doc.artboards[0];
    } else {
        var artboard = doc.artboards[1];
    }

    var totalImprints = ('0' + documentColors.length.toString()).slice(-2);

    artboard.name = artboard.name.substring(0, artboard.name.length-2) + totalImprints;
}

function writeDocumentColors() {
    doc.layers.getByName('info').groupItems.getByName('job').textFrames.getByName('meta').lines[1].contents = strDocumentColors;

    // setup counters
    var startChar = 0;
	var endChar = 0;

	for(var c = 0; c < documentColors.length; c++) {
		if (c > 0) {
			startChar = documentColors[c - 1].length + startChar + 1;
		}
		endChar = startChar + documentColors[c].length + 1;
		setDocumentColorStyle(documentColors[c], startChar, endChar);
	}
}

function setDocumentColorStyle(style, start, end) {
    var documentColorsText = doc.layers.getByName('info').groupItems.getByName('job').textFrames.getByName('meta').lines[1];

    for(var t = start; t < end; t++) {
        doc.characterStyles.getByName(style).applyTo(documentColorsText.characters[t], true);
    }
}

function setMetaDataPosition() {
    var mediaSize = 0;
    var offset = 0;

    if (documentColors.length == 1) {
        var artboard = doc.artboards[0];
        doc.artboards.setActiveArtboardIndex(0);
    } else {
        var artboard = doc.artboards[1];
        doc.artboards.setActiveArtboardIndex(1);
    }

    // reset Rulers
    doc.rulerOrigin = [0,0];
    doc.pageOrigin = [0,0];

    mediaSize = artboard.name.substring(0, artboard.name.length-5).slice(-1);

    if (mediaSize >= 3) {
        offset = 108;
    } else {
        offset = 36;
    }

    var jobMetadata = doc.layers.getByName('info').groupItems.getByName('Job');

    // reset Job Metadata position
    jobMetadata.left = 0;
    jobMetadata.top = jobMetadata.height;

    // position Job setMetadata
    jobMetadata.left = offset;
    jobMetadata.top = jobMetadata.top + offset;
}
