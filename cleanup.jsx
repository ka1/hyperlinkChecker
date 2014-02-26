app.scriptPreferences.version = 8; //Indesign CS6 and later. Use 7.5 to use CS5.5 features
app.scriptPreferences.userInteractionLevel = UserInteractionLevels.INTERACT_WITH_ALL; //INTERACT_WITH_ALL is the default anyway. just in case you want to change that.
app.scriptPreferences.enableRedraw = true;

var myDocument;
myDocument = app.documents.item(0);

//delete the layer
var helperLayer = returnLayerOrCreatenew("SourceLinkPageDiff");
helperLayer.remove();

//delete the object styles
var helperObjectStyle = returnObjectStyleOrCreatenew("SourceLinkPageDiffStyle");
helperObjectStyle.remove();
var helperTextframeObjectStyle = returnObjectStyleOrCreatenew("LinkPageDiffText");
helperTextframeObjectStyle.remove();

//delete the text styles
var helperTextframeParagraphStyle = returnParagraphStyleOrCreatenew("LinkPageDiff");
helperTextframeParagraphStyle.remove();

//delete colors
var color0 = returnColorOrCreateNew("DIFF-0", [60,0,100,0]);
var color1 = returnColorOrCreateNew("DIFF-1", [60,0,30,0]);
var color10 = returnColorOrCreateNew("DIFF-10", [0,45,100,0]);
var colorMore = returnColorOrCreateNew("DIFF-MORE", [0,100,100,0]);
var colorWeiss = returnColorOrCreateNew("DIFF-WEISS", [0,0,0,0]);

myDocument.swatches.item(color0.name).remove();
myDocument.swatches.item(color1.name).remove();
myDocument.swatches.item(color10.name).remove();
myDocument.swatches.item(colorMore.name).remove();
myDocument.swatches.item(colorWeiss.name).remove();








//Functions

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