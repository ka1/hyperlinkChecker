app.scriptPreferences.version = 8; //Indesign CS6 and later. Use 7.5 to use CS5.5 features
app.scriptPreferences.userInteractionLevel = UserInteractionLevels.INTERACT_WITH_ALL; //INTERACT_WITH_ALL is the default anyway. just in case you want to change that.
app.scriptPreferences.enableRedraw = true;

//Confirm each replacement of relative page references
var confirmReplacements = false;
var relativeReferencesOnSameSheet = false; //true if "gegenüberliegende Seite" should be written
var relativeReferencesWithin1Page = false; //true if page references to next or previous page should be written

var helperLayer;
var myDocument;
myDocument = app.documents.item(0);

//safe active layer
var savedActiveLayer = myDocument.activeLayer;

helperLayer = returnLayerOrCreatenew("SourceLinkPageDiff");
helperLayer.remove();
helperLayer = returnLayerOrCreatenew("SourceLinkPageDiff");

var color0, color1, color10, colorMore, colorWeiss, colorNull;

color0 = returnColorOrCreateNew("DIFF-0", [60,0,100,0]);
color1 = returnColorOrCreateNew("DIFF-1", [60,0,30,0]);
color10 = returnColorOrCreateNew("DIFF-10", [0,45,100,0]);
colorMore = returnColorOrCreateNew("DIFF-MORE", [0,100,100,0]);
colorWeiss = returnColorOrCreateNew("DIFF-WEISS", [0,0,0,0]);
colorNull = returnColorOrCreateNew("DIFF-NULL", [0,0,0,50]);

var count0 = 0, count1 = 0, count10 = 0, countMore = 0, countNull = 0;

var helperTextframeParagraphStyle = returnParagraphStyleOrCreatenew("LinkPageDiff",null,{
	pointSize:10,
	appliedFont: "Arial",
	fontStyle: "Bold",
	fillColor: colorWeiss.name
});

var helperTextframeObjectStyle = returnObjectStyleOrCreatenew("LinkPageDiffText",{
	enableParagraphStyle: true,
	appliedParagraphStyle: helperTextframeParagraphStyle,
	enableStroke: true,
	strokeColor: myDocument.swatches[0],
	enableFill: true,
	fillColor: myDocument.swatches[0],
	enableTextFrameGeneralOptions: true,
	textFramePreferences: {
		useNoLineBreaksForAutoSizing: true,
		textColumnCount: 1,
		autoSizingType: AutoSizingTypeEnum.HEIGHT_AND_WIDTH,
		autoSizingReferencePoint: AutoSizingReferenceEnum.CENTER_POINT
	},
	enableTextFrameAutoSizingOptions: true,
	objectEffectsEnablingSettings: {
		enableTransparency: true },
	transparencySettings: {
		blendingSettings: {
			blendMode: BlendMode.NORMAL, opacity: 70}},
});


var helperObjectStyle = returnObjectStyleOrCreatenew("SourceLinkPageDiffStyle",{
	enableFill: true,
	fillColor: myDocument.swatches[5],
	objectEffectsEnablingSettings: {
		enableTransparency: true },
	transparencySettings: {
		blendingSettings: {
			blendMode: BlendMode.NORMAL, opacity: 70}},
	enableStroke: true,
	strokeColor: myDocument.swatches[0],
	enableTextWrapAndOthers: true,
	textWrapPreferences: {
		textWrapMode: TextWrapModes.NONE}
});

if (confirm("Update (= reset) all links before processing?",true,"Update links")){
	for(var c = 0; c < myDocument.crossReferenceSources.length; c++){
		myDocument.crossReferenceSources.item(c).update();
	}
}

for(i = 0; i < myDocument.hyperlinks.length; i++){
	var currentHyperlink = myDocument.hyperlinks[i];
	var currentDestination = currentHyperlink.destination;
	var currentSource = currentHyperlink.source;

	if (currentSource instanceof CrossReferenceSource == false){
		continue;
	}
	
	//only search for Bildunterschrift links and only search for the link style "* Seitenzahl"
	if (currentDestination.destinationText.appliedParagraphStyle.name == 'Bildunterschrift' && currentSource.appliedFormat.name.match(/.*Seitenzahl/)){
		var currentSourceText = currentSource.sourceText;
		var currentSourcePage = currentSourceText.insertionPoints.firstItem().parentTextFrames[0].parentPage;
		var currentDestinationPage = currentDestination.destinationText.insertionPoints.firstItem().parentTextFrames[0].parentPage;
		
		var margin = 5;
		var ry1 = currentSourceText.insertionPoints.item(0).baseline;
		var ry2 = currentSourceText.insertionPoints.item(0).baseline - currentSourceText.texts[0].ascent;
		var rx1 = currentSourceText.insertionPoints.item(0).horizontalOffset;
		var rx2 = currentSourceText.insertionPoints.item(-1).endHorizontalOffset;

		//the rectangle that contains the color mark
		var markingRectangle = currentSourcePage.rectangles.add({geometricBounds:[ry1 + margin,rx1 - margin,ry2 - margin,rx2 + margin], label: "pageDifferenceIdentifier",itemLayer: helperLayer, appliedObjectStyle: helperObjectStyle});
		//the textframe that contains the page difference
		var currentIdentifierTextframe = currentSourcePage.textFrames.add(helperLayer, LocationOptions.UNKNOWN,{geometricBounds:[ry1,rx1,ry2,rx2], label: "pageDifferenceIdentifierText",appliedObjectStyle: helperTextframeObjectStyle});

		var pageDiff ;
		if (currentDestinationPage == null){
			//handle references to invalid pages
			pageDiff = null;
		} else {
			//positive if links to page before source page, negative, if links to page after source page
			pageDiff = parseInt(currentSourcePage.name) - parseInt(currentDestinationPage.name);
		}
	
		if (pageDiff == null){
			$.writeln("null");
			currentIdentifierTextframe.contents = "Invalid page";
		}
		else if (pageDiff == 0 || Math.abs(pageDiff) == 1){
			//currentSource.showSource();
			//alert('Detected same page reference');
			
			//add relative page reference
			//confirm this? (is globally overwritten by confirmReplacements
			var doConfirm = true;
			if (pageDiff == 0) {
				checkAndDeletePageReference(currentSourceText);
				currentIdentifierTextframe.contents = "Diese Seite";
				markingRectangle.fillColor = color0.name;
				//nothing to add, just delete the text
				doConfirm = false;
			} else if (pageDiff == 1) {
				currentIdentifierTextframe.contents = "vorherige Seite";
				markingRectangle.fillColor = color1.name;
				if (currentSourcePage.side == PageSideOptions.RIGHT_HAND){
					doConfirm = addRelativeReferenceSameSheet(currentSourceText,", gegenüberliegende Seite");
				} else {
					doConfirm = addRelativeReference1Page(currentSourceText,", vorherige Seite");
				}
			} else {
				//if pageDiff == -1
				currentIdentifierTextframe.contents = "nächste Seite";
				markingRectangle.fillColor = color1.name;
				if (currentSourcePage.side == PageSideOptions.LEFT_HAND) {
					doConfirm = addRelativeReferenceSameSheet(currentSourceText,", gegenüberliegende Seite");
				} else {
					doConfirm = addRelativeReference1Page(currentSourceText,", nächste Seite");
				}
			}
			
			//confirm changes
			if (confirmReplacements && doConfirm){
				helperLayer.visible = false;
				currentSource.showSource();
				if (!confirm("Änderung so OK (Ja) oder rückgängig machen (nein)?", true, "Änderung prüfen")){
					//undo changes
					currentSource.update();
				}
				helperLayer.visible = true;
			}
			
		} else if (pageDiff > 0){
			currentIdentifierTextframe.contents = Math.abs(pageDiff )+ " Seiten vorher";
		} else {
			currentIdentifierTextframe.contents = Math.abs(pageDiff) + " Seiten nachher";
		}
	
		//color the rectangle
		if (pageDiff == null){
			markingRectangle.fillColor = colorNull.name;
			countNull++;
		}
		else if (pageDiff == 0){
			count0++;
		} else if (Math.abs(pageDiff ) == 1){
			count1++;
		} else if (Math.abs(pageDiff ) < 10){
			markingRectangle.fillColor = color10.name;
			count10++;
		} else if (Math.abs(pageDiff) >= 10){
			markingRectangle.fillColor = colorMore.name;
			countMore++;
		}
	}
}

//restore active layer
myDocument.activeLayer = savedActiveLayer ;

alert("Ready.\nSame page: " + count0 + "\nWithin 1 page: " + count1 + "\nWithin 9 pages: " + count10 + "\nMore than 9 away: " + countMore + "\nInvalid: " + countNull);






//Functions

function checkAndDeletePageReference(currentSourceText){
	//delete page number and ", Seite "
	//see if it ends with a page number AND if the contents contains the SEITE text
	if (currentSourceText.characters.item(-1).textVariableInstances.length == 1
	&& currentSourceText.characters.item(-1).textVariableInstances.item(0).associatedTextVariable.name == 'TV XRefPageNumber'
	&& currentSourceText.contents.match(/.*, Seite .$/)){
		//delete the last character
		currentSourceText.characters.item(-1).remove();
		currentSourceText.contents = currentSourceText.contents.replace(/, Seite /,"");
		return true;
	}
	else {
		return false; //so that other functions know that nothing happened / nothing was deleted
	}
}

//changes the source text and returns true, so that the text can be confirmed (in confirmations are on)
//or leaves everything as is (if relativeReferencesWithin1Page is false) and returns false
function addRelativeReference1Page(currentSourceText,text){
	if (relativeReferencesWithin1Page){
		if(checkAndDeletePageReference(currentSourceText)){
			currentSourceText.contents += text;
			return true;
		}
		else {
			return false;
		}
	} else {
		return false;
	}
}

//similar to addRelativeReference1Page
function addRelativeReferenceSameSheet(currentSourceText, text){
	//always delete absolute page reference
	if (checkAndDeletePageReference(currentSourceText) && relativeReferencesOnSameSheet){
		currentSourceText.contents += text;
		return true;
	} else {
		return false;
	}
}

function returnColorOrCreateNew(colorName, cmykValues){
	var color;
	color = myDocument.colors.item(colorName);
	try{
		name = color.name;
	} catch (e) {
		color = myDocument.colors.add({
			name: colorName,
			model: ColorModel.PROCESS,
			space: ColorSpace.CMYK,
			colorValue: cmykValues
		});
	}

	return color;
}

function returnLayerOrCreatenew(layerName){
	var layer;
	layer = myDocument.layers.item(layerName);
	try{
		name = layer.name;
	} catch(e) {
		layer = myDocument.layers.add({name: layerName});
	}
	return layer;
}

//return an existing object style or create a new object style. if preferences are given, these will be applied to the new object style only.
function returnObjectStyleOrCreatenew(objectStyleName, newObjectStylePreferences){
	var objectStyle;
	objectStyle = myDocument.objectStyles.item(objectStyleName);
	try{
		name = objectStyle.name;
	} catch(e) {
		if (newObjectStylePreferences != null) {
			newObjectStylePreferences.name = objectStyleName;
			objectStyle = myDocument.objectStyles.add(newObjectStylePreferences);
		} else {
			objectStyle = myDocument.objectStyles.add({name: objectStyleName});
		}
	}
	return objectStyle;
}

//return the reference to a paragraph style. if the style did not exist, create the style
function returnParagraphStyleOrCreatenew(stylename, groupname, stylePreferences){
	var style;
	var group;

	//prepare style preferences. if none are given, only include the name
	if (stylePreferences == null){
		stylePreferences = {name: stylename};
	} else {
		stylePreferences.name = stylename;
	}
	
	if (!groupname){ //is the style in a group?
		style = myDocument.paragraphStyles.item(stylename);
	} else {
		try { //add group first, if it does not exist
			group = myDocument.paragraphStyleGroups.itemByName(groupname);
			gname = group.name; //will trigger error if group does not exist
		}
		catch(e){
			group = myDocument.paragraphStyleGroups.add({name: groupname});
			style = group.paragraphStyles.add(stylePreferences); //then add style in the group
		}
		style = group.paragraphStyles.itemByName(stylename); //select style in group (for the second time, if the style had already been created, but that should be ok)
	}
	try{
		name=style.name; //will trigger error if style does not exist
	} catch(e) {
		if (group != null){
			style = group.paragraphStyles.add(stylePreferences);
		} else {
			style = myDocument.paragraphStyles.add(stylePreferences);
		}
	}
	return style;	
}