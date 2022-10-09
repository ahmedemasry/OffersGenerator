   
   
   //Defining Constants & Strings
   const BGLayerName = "BG";
   const UI_IMPORTCSV = "Import CSV file";
   const UI_SAVENOTE = "IMPORTANT NOTE: We will save this PSD fle to run the script. \nIs it ok?";
   const UI_FIRSTCOLUMN_WIDTH = 120;
   const TEXT_LAYER_DIRECTIVE = "#";
   const VISIBLE_LAYER_DIRECTIVE = "*";
   const FILE_NAME_TAG = "File_Name";
   const LAYER_AR_TAG = "(ar)";
   const LAYER_EN_TAG = "(en)";
   const arSuffix = "_AR";
   const enSuffix = "_EN";
   const sizeSuffix = "";
   const IdTag = "Artboard";
 
   
   const doc = app.activeDocument;
   const artboards = doc.layers;
   const bgArtboard;
   const bgLayers;
   
   try{
       bgArtboard = artboards[BGLayerName];
       bgLayers = bgArtboard.layers;
   }catch(e){
       alert("Make sure you have an artboard with the name 'BG'.", "'BG' Artboard !", true);
   }
    
   
   //Defining Essential Variables
   var savePath;
   var moveToLayerID;
      var headers = [];
   var directives = [];
   
      
   //csvFile & csvDaya will be inialized from startDialog(run);
   var csvFile;
   var csvData;
        savePath = app.activeDocument.path;
        moveToLayerID = bgLayers.length;
        startDialog(run);
   
   
   function run(){
       try{
            //Calling the generateOffers() function within the suspendHistory to save all steps done only in one history element; to easily undo any unexpected errors.
           doc.suspendHistory("Generate Offers", "generateOffers();");
        }
        catch(e){
            
            alert("Error, Please Double Check Naming (Hint: Check Upper and Lower case and Layers Existance). \n\n " + e, "Error!", true);
            resetFile();
        }
   }
   
   
   //------------------------------------------------------------------------------------
   
   function generateOffers(){
        
       //Iterating over all Artboards and Copy its content to BG
       for(var i = 0; i < csvData.length; i++){

            var movedLayers = [];
            var corrsepondingCsvIndex = i;
            var currentCsvRow = csvData[corrsepondingCsvIndex];
            var offerName = currentCsvRow[FILE_NAME_TAG.toLowerCase()];
            var currentArtboard;
            var offerLayers;
            try{
                currentArtboard =  artboards.getByName(currentCsvRow[IdTag.toLowerCase()]);
                offerLayers = currentArtboard.layers;
            }catch(e){
                //Handling the mistakes in naming artboards and sheet columns             
                if(currentCsvRow[IdTag.toLowerCase()] === ""){
                    continue;
                }
            
                offerLayers = [];
            }

             //Iterating over all Layers in each Artboard 
            for(var j = 0; j < offerLayers.length && offerLayers.length>0;){

                currentLayer = offerLayers[j];
               
                if(currentLayer.name === "Layer 0" || currentLayer.visible === false){
                    j++;
                    continue;
                }

                //If this layer is visible, move it to the BG artboard
                var fromID = currentLayer.id;
                var toID = moveToLayerID;
                doc.activeLayer = currentLayer;
                moveLayerToLayerSet(fromID, toID);
                movedLayers.push(currentLayer);

            }
        
            //Disable All Artboards except for the BG
             visibleLayersExceptBG(false);
            
            //For Each Offer, Prepare the file to save
                
                //Setting Layers Using Directives
                for(var d=0; d<directives.length; d++){
                    
                    var layerName = headers[d];
                    //If it is Text Layer, Use text in the element.
                    if(directives[d] === TEXT_LAYER_DIRECTIVE){
                        var textLayer = bgLayers.getByName(layerName);
                        textLayer.textItem.contents = currentCsvRow[layerName];
                    }
                    //If it is Visible Directive, Make it invisible if value = 0, else visible.                    
                    else if(directives[d] === VISIBLE_LAYER_DIRECTIVE){
                        var layer = bgLayers.getByName(layerName);
                        if(currentCsvRow[layerName] === '0'){
                            layer.visible = false;
                        }
                        else{
                            layer.visible = true;
                        }
                    }
                    
                 }
                
                //Arabic Saving
                var arText = visibileTextArEnTag(true, currentCsvRow);
                if(arText){
                    saveJpeg (offerName + arSuffix + sizeSuffix);
                }
                        
                //English Saving
                var enText = visibileTextArEnTag(false, currentCsvRow);
                if(enText){
                    saveJpeg (offerName + enSuffix + sizeSuffix);
                }
                
                //In case of one language.
                if(!arText && !enText){
                    saveJpeg (offerName + sizeSuffix);
                }

            //Show All Artboards except for the BG
            visibleLayersExceptBG(true);
            
            
    //TODO Make it reset manually by moving layers back OR removing layers (File already is reset in case of errors)-----------------------------------------------------------
    //Improved from (252) To (130) seconds for a test case with about 194% in speed increas. 48% of time saved
    //Improved from (182) To (104) seconds for a test case with about 175% in speed increas. 43% of time saved
    //Improved from (14) To (10) seconds for a test case with about 140% in speed increase 28% of time saved
    // ============  On Average 170% in speed increase 40% of time saved
            //Reset File
            visibleLayersExceptBG(true);
            for(var n = 0; n < movedLayers.length; n++){
                currentLayer = movedLayers[n];
                currentLayer.remove();
            }
       }
       resetFile();
       alert("All Done!");
    }


//==================================================================================================
//Utility Functions
function visibileTextArEnTag(bool_AR, currentCsvRow){
    var languageLayersCount = 0;
    for(var a = 0; a < bgLayers.length; a++){
        var layerName = bgLayers[a].name;
        var isArabicLayer = (layerName).toLowerCase().indexOf(LAYER_AR_TAG) !== -1;
        var isEnglishLayer = (layerName).toLowerCase().indexOf(LAYER_EN_TAG) !== -1;
        //Skip Invisivle Layers (Forced By Sheet)
        if(currentCsvRow[layerName] !== undefined){
            if(currentCsvRow[layerName] === '0'){
                bgLayers[a].visible = false;
                continue;
            }
        }
        
        if(isArabicLayer){
            bgLayers[a].visible = bool_AR;
            languageLayersCount += bool_AR;
        }else if (isEnglishLayer){
              bgLayers[a].visible = !bool_AR;
              languageLayersCount += !bool_AR;
        }
    }

    //Return the Boolean Value
    return !(!languageLayersCount);
}

function loadCSVFile(csvFile){

    csvFile.open('r');
    var str = csvFile.read();
    csvFile.close();
    
    return csvToArray(str);
}

function csvToArray(str) {
       var delimiter = ",";
      // slice from start of text to the first \n index
      // use split to create an array from string by delimiter
      headers = str.slice(0, str.indexOf("\n")).split(delimiter);
      
      for(var i = 0; i < headers.length; i++){
          if(headers[i].slice (0, 1) === TEXT_LAYER_DIRECTIVE){
              directives[i] = TEXT_LAYER_DIRECTIVE;
              headers[i] = headers[i].slice (1); 
          }
          else  if(headers[i].slice (0, 1) === VISIBLE_LAYER_DIRECTIVE){
              directives[i] = VISIBLE_LAYER_DIRECTIVE;
              headers[i] = headers[i].slice (1); 
          }
          else {
              directives[i] = "";
              headers[i] = headers[i].toLowerCase(); 
          }
      }

      // slice from \n index + 1 to the end of the text
      // use split to create an array of each csv value row
      const rows = str.slice(str.indexOf("\n") + 1).split("\n");
      // Map the rows
      // split values from each row into an array
      // use headers.reduce to create an object
      // object properties derived from headers:values
      // the object passed as an element of the array
      const arr = [];
      for(var i = 0; i < rows.length; i++){
          var row = rows[i];
          var values = row.split(delimiter);
          var obj = {};
          var hIndex = 0
           for(var j = 0; j < values.length; j++){
               var k = j;
               //For values having comma ','
               if( values[j].slice (0, 1) === '\"'){
                   values[j] = values[j].slice (1, values[j].length) + ",";
                   while(values[++k].slice(-1, values[k].length) !=  '\"'){
                       values[j] += values[k] + ",";
                   }
                    values[j] += values[k].slice(0, -1);
               }
               if(((values[j]).toLowerCase().indexOf("/") !== -1 || (values[j]).toLowerCase().indexOf("\\") !== -1 ) && headers[hIndex] === FILE_NAME_TAG.toLowerCase()){
                //Handling unsupported file names
                   values[j] = values[j].replace(/[^a-zA-Z0-9 ]/g, '');
               }
               obj[headers[hIndex]] = values[j];
               hIndex++;
               j = k;
           }
          arr.push (obj);
      }

  
      // return the array
      return arr;
    }



function startDialog(functionCallBack){
    //The main options for importing the .CSV file, choosing the place to save ouptut files and selecting the layer to inject the components above
     var dlgMain = new Window("dialog", UI_IMPORTCSV);

	dlgMain.orientation = 'column';
	dlgMain.alignChildren = 'left';

// -- top of the dialog, first line
    dlgMain.add("statictext", undefined, UI_SAVENOTE);

// -- two groups, one for left and one for right ok, cancel
	dlgMain.grpTop = dlgMain.add("group");
	dlgMain.grpTop.orientation = 'row';
	dlgMain.grpTop.alignChildren = 'top';
	dlgMain.grpTop.alignment = 'fill';

	// -- group top left 
	dlgMain.grpTopLeft = dlgMain.grpTop.add("group");
	dlgMain.grpTopLeft.orientation = 'column';
	dlgMain.grpTopLeft.alignChildren = 'left';
	dlgMain.grpTopLeft.alignment = 'fill';
    
	// -- the second line in the dialog
	dlgMain.grpSecondLine = dlgMain.grpTopLeft.add("group");
	dlgMain.grpSecondLine.orientation = 'row';
	dlgMain.grpSecondLine.alignChildren = 'center';

    dlgMain.etDestination = dlgMain.grpSecondLine.add("statictext", undefined, UI_IMPORTCSV);
    dlgMain.etDestination.preferredSize.width = UI_FIRSTCOLUMN_WIDTH;

    dlgMain.btnImportCsv = dlgMain.grpSecondLine.add("button", undefined, "Import CSV...");
  //--------------------------------------------------------------------------------------------------------------------------------  
    dlgMain.btnImportCsv.onClick = function() {
        csvFile = File.openDialog('Select File', "*.csv");
        if(csvFile === null) return;
        if(csvFile.name.split('.').pop().toLowerCase() === "csv"){
            dlgMain.etDestination.text = csvFile.name;
        }
	}
  //-------------------------------------------------------------------------------------------------------------------------------- 
	// -- the third line in the dialog
    dlgMain.grpThirdLine = dlgMain.grpTopLeft.add("group");
    dlgMain.section3Text = dlgMain.grpThirdLine.add("statictext", undefined, "Save Files At");
    dlgMain.section3Text.preferredSize.width = UI_FIRSTCOLUMN_WIDTH;
     dlgMain.btnBrowse = dlgMain.grpThirdLine.add("button", undefined, "Browse...");
     dlgMain.section3Text.text = savePath;
     dlgMain.btnBrowse.onClick = function() {
        var defaultFolder = savePath;

		var testFolder = new Folder(savePath);

		if (!testFolder.exists) {

			defaultFolder = "~";

		}

		var selFolder = Folder.selectDialog("Select Destination", savePath);

		if ( selFolder != null ) {

	       savePath = selFolder.fsName;
            dlgMain.section3Text.text = savePath;

	    }

		dlgMain.defaultElement.active = true;
	}

  //-------------------------------------------------------------------------------------------------------------------------------- 
	// -- the 4th line in the dialog
    dlgMain.grpThirdLine = dlgMain.grpTopLeft.add("group");

   
    

    const max = bgLayers.length-1;
    var layer = 0;
    moveToLayerID = getLayerIndexByID(bgLayers[max - layer].id);
    moveToLayerName = bgLayers[max - layer].name;
    
    dlgMain.layerText = dlgMain.grpThirdLine.add("statictext", undefined, "Layer: ");
    dlgMain.slider = dlgMain.grpThirdLine.add("scrollbar", undefined, layer, 0, max);
    dlgMain.slider.stepdelta = 1.000001;
    dlgMain.slider.helpTip = "Insert Above Layer: " + moveToLayerName +".";
    dlgMain.slider.preferredSize.width = UI_FIRSTCOLUMN_WIDTH;
    dlgMain.layerValueText = dlgMain.grpThirdLine.add("statictext", undefined, "Above: " + moveToLayerName);
    dlgMain.layerValueText.preferredSize.width = 100;
    
    dlgMain.slider.onChanging = function(){
        layer = parseInt(dlgMain.slider.value);
        dlgMain.slider.helpTip = "Insert Above Layer: " + bgLayers[max - layer].name +".";
        moveToLayerID = getLayerIndexByID(bgLayers[max - layer].id);
        dlgMain.layerValueText.text = "Above: " + bgLayers[max - layer].name ;
        
    }
    
    //--------------------------------------------------------------------------------------------------------------------------------   
    dlgMain.grpTopLeft.add("statictext", undefined, "");
    dlgMain.grpTopLeft.add("statictext", undefined, "Ⓒ Ahmed M. ElMasry, 29/6/2022");
    dlgMain.grpTopLeft.add("statictext", undefined, " ahmedemasry@gmail.com");
    
    
    
  //=================RIGHT==========================================================
	// the right side of the dialog, the ok and cancel buttons
	dlgMain.grpTopRight = dlgMain.grpTop.add("group");
	dlgMain.grpTopRight.orientation = 'column';
	dlgMain.grpTopRight.alignChildren = 'fill';
    
    // Save to Apply current Layout  
	dlgMain.btnRun = dlgMain.grpTopRight.add("button", undefined, "Save and Run" );
    dlgMain.btnRun.onClick = function() {
        if(csvFile === undefined || csvFile === null){alert("Please Select A CSV File First!.");}
        else if(csvFile.name.split ('.').pop().toLowerCase() !== "csv"){
            alert("Please Select A CSV File First!.");
        }
        else{
            saveFile (); 
            csvData = loadCSVFile(csvFile);
            dlgMain.close();            
            functionCallBack();
            
        }
	}

	dlgMain.btnCancel = dlgMain.grpTopRight.add("button", undefined, "Cancel" );
    dlgMain.btnCancel.onClick = function() { 
		dlgMain.close(); 
	}

	dlgMain.defaultElement = dlgMain.btnRun;
	dlgMain.cancelElement = dlgMain.btnCancel;
    
    dlgMain.center();
    dlgMain.show();


}

function saveJpeg(offerName){
    
    var file = new File(savePath+ "/" +offerName+".jpg");
      saveOptions = new ExportOptionsSaveForWeb();
      saveOptions.format = SaveDocumentType.JPEG
      saveOptions.optimized = false;
      saveOptions.quality = 95;
      doc.exportDocument(file, ExportType.SAVEFORWEB, saveOptions);
    
    
}

function visibleLayersExceptBG(visibility){
//Hide/Show All Artboards except for the BG
   for(var i = 0; i < artboards.length; i++){
       currentArtboard =  artboards[i];
       if(currentArtboard === bgArtboard){ 
           continue;
       }
        currentArtboard.visible = visibility;
   }
}

function alertSaveFile(functionCallBack){
    var g = new Window('dialog', 'Save File?');
    g.add('statictext', undefined, "We Will Save this File to run this script. \nIs it ok?");
    buttonOk = g.add('button', undefined, 'Save');
    buttonCancel = g.add('button', undefined, 'Cancel');
    buttonOk.onClick = function() {
         saveFile ();   
         g.close();
         functionCallBack();
     }
    buttonCancel.onClick  = function() {
        g.close();
        alert("The script will be canceled, You Have to Save.");
     }

    g.center(), g.show();
}
     
     
function saveFile(){
    var idsave = charIDToTypeID( "save" );
    var desc1844 = new ActionDescriptor();
    var idIn = charIDToTypeID( "In  " );
    desc1844.putPath( idIn, new File( doc.path+"/"+doc.name) );
    var idDocI = charIDToTypeID( "DocI" );
    desc1844.putInteger( idDocI, 542 );
    var idsaveStage = stringIDToTypeID( "saveStage" );
    var idsaveStageType = stringIDToTypeID( "saveStageType" );
    var idsaveSucceeded = stringIDToTypeID( "saveSucceeded" );
    desc1844.putEnumerated( idsaveStage, idsaveStageType, idsaveSucceeded );
    executeAction( idsave, desc1844, DialogModes.NO );
}
function resetFile(){
    var idRvrt = charIDToTypeID( "Rvrt" );
    executeAction( idRvrt, undefined, DialogModes.NO );
}
function saveArtboardToFile(){
    
    var idAdobeScriptAutomationScripts = stringIDToTypeID( "AdobeScriptAutomation Scripts" );
    var desc1459 = new ActionDescriptor();
    var idjsNm = charIDToTypeID( "jsNm" );
    desc1459.putString( idjsNm, """Artboards to Files...""" );
    var idjsMs = charIDToTypeID( "jsMs" );
    desc1459.putString( idjsMs, """undefined""" );
    executeAction( idAdobeScriptAutomationScripts, desc1459, DialogModes.NO );

}


function moving(v1, v2, v3) {

     function sTT(v) {return stringIDToTypeID(v)} eval('(ref1 = eval(AR = "new ActionReference()")).put'  + (!v1 ? 'Enumerated'  : ((iN = isNaN(v1)) ? 'Name' : 'Index' ))

     + '(l = sTT("layer"), ' + (!v1 ? 'sTT("ordinal"), sTT("targetEnum")' : 'v1'  + (!iN ? -1 : "")) + ')'); (dsc1 = new ActionDescriptor()).putReference(sTT('null'), ref1)

     if (v2) eval('(ref = eval(AR)).put' + ((iN = isNaN(v2)) ? 'Name' : 'Index' ) + '(l, v2 + (!iN ? -1 : ""))'); (ref2 = eval(AR)).putIndex(l, v2 || (v2 != undefined && iN) ?

     executeActionGet(ref).getInteger(sTT('itemIndex')) - (v3 || 0) : (v2 || ~~(lyr = activeDocument.layers)[lyr.length - 1].isBackgroundLayer)), dsc1.putReference(sTT('to'), ref2)

     dsc1.putInteger(sTT('version'), 5), dsc1.putBoolean(sTT('adjustment'), false), $.level = 0; try{executeAction(sTT('move'), dsc1, DialogModes.NO)} catch(err){}
}







function moveLayerToLayerSet( fromID, toID ){
    var newToID = toID;

    var idmove = charIDToTypeID( "move" );
    var desc = new ActionDescriptor();
    var idnull = charIDToTypeID( "null" );
        var ref131 = new ActionReference();
        var idLyr = charIDToTypeID( "Lyr " );
        var idOrdn = charIDToTypeID( "Ordn" );
        var idTrgt = charIDToTypeID( "Trgt" );
        ref131.putEnumerated( idLyr, idOrdn, idTrgt );
        ref131.putIdentifier( idLyr , fromID);
    desc.putReference( idnull, ref131 );
    var idT = charIDToTypeID( "T   " );
        var ref132 = new ActionReference();
        var idLyr = charIDToTypeID( "Lyr " );
        ref132.putIndex( idLyr, newToID );
    desc.putReference( idT, ref132 );
    var idAdjs = charIDToTypeID( "Adjs" );
    desc.putBoolean( idAdjs, false );
    var idVrsn = charIDToTypeID( "Vrsn" );
    desc.putInteger( idVrsn, 5 );
    var idLyrI = charIDToTypeID( "LyrI" );
        var list104 = new ActionList();
        list104.putInteger(fromID );
    desc.putList( idLyrI, list104 );
    
   
   try{
    executeAction( idmove, desc, DialogModes.NO );
    }catch(e){alert(e);}

};


function getLayerIndexByID(ID){

var ref = new ActionReference();

ref.putIdentifier( charIDToTypeID('Lyr '), ID );

try{ 

activeDocument.backgroundLayer; 

return executeActionGet(ref).getInteger(charIDToTypeID( "ItmI" ))-1; 

}catch(e){ 

return executeActionGet(ref).getInteger(charIDToTypeID( "ItmI" )); 

}

};
   
   
   
 
 
 
 
 
 

 
 
 
 
 //======================================================================================================================================
 //======================================================================================================================================
 //======================================================================================================================================
 //======================================================================================================================================
 
 