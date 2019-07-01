// PREPRINT BATCHING SCRIPT FOR ADOBE ILLUSTRATOR

var abc                 = 0;
var allPDFFiles         = [];

 // FILES & FOLDERS
 var rootString         = 'C:\\Users\\peter\\Desktop\\OJ\\Test for Peter - Copy';
 var processedFolder    = Folder( rootString + '\\Processed' );
 var oneupsFolder       = Folder( rootString + '\\Oneups' );
 var savedFolder        = Folder( rootString + '\\Saved' );
 var templatesFolder    = Folder( rootString + '\\Templates');

// ****************************************
// THIS IS WHERE YOU PARAMETERS ARE LOADED.
// WITH THE FIRST ROW AS AN EXAMPLE:
// ******************************************************************
//  ROW 1 -> [ [ 'imtb' , 'imbb' , 'imzz' , 'imbvb' , 'imdd' ] , 90 ] ,   
// FIRST ELEMENT ^^^^^^ IS THE STRING USED TO FIND DYELINE TEMPLATE.
// THE FOLLOWING ELEMENTS, IF THERE ARE ANY, WILL GRAB THE FIRST ELEMENT AS DYE TEMPLATE.
// I.E.- IF 'imzz' IS CALLED IN THE SCRIPT, IT WILL GRAB THE STRING 'imtb' AND 
      // GENERATE ALL NECESSARY STRINGS.
// THE NUMBER AT THE END OF THE ROW IS THE ROTATION APPLIED
// IF YOU HAVE A NEW ONEUP THAT NEEDS TO USE 'PSS' (AND ITS 90 DEGREE ROTATION) 
      // YOU WOULD PLACE IT IN THE ROW WHERE PSS IS THE FIRST ELEMENT [2ND ROW].
var prodFiles = [// DYELINE \/
                    [ [ 'imtb' , 'imbb' , 'imzz' , 'imbvb' , 'imdd' ]   ,90 ] ,   
                    [ [ 'pss' , 'ptt' , 'ptz' , 'pppp' , 'pooped' ]     , 0 ] ,
                    [ [ 'usalb' , 'imibib' , 'imabab' ]                 , 0 ] ,
                    [ [ 'hsr10' ]                                       , 0 ] ,
                    [ [ 'lb60' ]                                        , 0 ] ,
                    [ [ 'lo30' ]                                        , 0 ] ,
                    [ [ 'hsr20' ]                                       , 0 ] ,
                    [ [ 'sss2' ]                                        , 0 ] ,
                    [ [ 'w12count' ]                                    , 0 ] ,
                    [ [ 'hs1z' ]                                        , 0 ] ,
                    [ [ 'w2ep' ]                                        , 0 ] ,
                 // [ [ 'hsr10' ]                                       , 0 ] ,
                    [ [ 'sss8' ]                                        , 0 ] ,
                    [ [ 'es2' ]                                         , 0 ] ,
                    [ [ 'hsc1' ]                                        ,90 ] ,
                    [ [ '2ep' , 'dubcup' , 'hs1z' ]                     , 0 ] ,
                    [ [ '2in' , 'hssk' , 'lb450' , 'lighter' ]          , 0 ] ,
                    [ [ 'ed8' , 'lmroll' , 'mft' , 'tr8901' ]           , 0 ]
                ];

// 1. GET PDF FILES FROM ONEUPS FOLDER
function getOneups ( startFolder ) {
    var folderContents;
    folderContents = startFolder.getFiles( );
    for ( var y = 0; y < folderContents.length; y ++ ) {
        if   ( folderContents[ y ].toLocaleString() == "[object Folder]" ) {
            getOneups ( Folder( folderContents [ y ].path + "//" + folderContents [ y ].name ) );
        }
        if ( folderContents[ y ].name.match( /(\.pdf)+/ ) ) {
            allPDFFiles.push( folderContents[ y ].fullName ); } }
    iteratePDFs() }

// *******************
function iteratePDFs() {
    for ( var i  = 0; i < allPDFFiles.length; i ++ ) { 
        getTemplate( allPDFFiles[ i ] );
        abc ++ } }

// *******************************
function getTemplate( fileToMatch ) {
    for ( var i = 0; i < prodFiles.length; i ++ ) {
        var ouSubarray = prodFiles[ i ][ 0 ];
        for ( var j = 0; j < ouSubarray.length; j ++ ) {
            var reggie = new RegExp( ouSubarray[ j ].toUpperCase( ) )
            if ( fileToMatch.match( reggie ) ) {
                var degs = prodFiles[ i ][ 1 ];
                var templateGot = File( templatesFolder + '\\Dieline_' + ouSubarray[ 0 ].toUpperCase() + '.pdf' );
                placerDoer( templateGot, fileToMatch, degs );
            } } } }

// ********************************************
function placerDoer( tempFi, fileToPlace, degs ) {
    app.open( tempFi );
    layerUnlocker();
    var oneupsFolderString = rootString + '\\Oneups';
    var n = oneupsFolderString.indexOf( '\\Oneups' );
    var fTPSubstring = fileToPlace.substring( n - 5 ).replace( '.pdf', '' );
    var savedFileString = rootString + '\\Saved\\IMP_' + fTPSubstring + '.pdf';
    var processedOneup = File( fileToPlace );
    processedOneup.copy( processedFolder + '\\' + fTPSubstring + '_PROCSD.pdf' );
    var oneupFileForGroup = File( fileToPlace );
    zamena ( savedFileString, oneupFileForGroup, degs );
    processedOneup.remove( );
}

// ********************
// UNLOCK LAYERS
function layerUnlocker( ) {
    // DECIDE TEMPLATE HERE, THEN RUN APPROPRIATE 
    // ACTIONS DEFINED ABOVE ON TEMPLATE
    if ( app.documents.length > 0 ) {
        var actDoc = app.activeDocument;
        var layersInDoc = actDoc.layers;
        for ( i = 0; i < layersInDoc.length; i++ ) {
            layersInDoc[i].locked = false;
        } } }

// ******************************************
// ZAMENA
function zamena( sav, oneupFileForGroup, degs ) { 
    var activeDocu = app.activeDocument;
    var uig = activeDocu.placedItems.add();
    uig.file = oneupFileForGroup;
    app.activeDocument.selectObjectsOnActiveArtboard();
    uig.selected = true;
    var mySelection = activeDocument.selection;
    if ( mySelection.length > 0 ) {
        if ( mySelection instanceof Array ) {
            var goal;
            goal = mySelection[0]; 
            var degs = degs;
            // ****** CALLBACK
            goal.rotate( degs );
            // GOAL IS THE PLACED ONEUP IMAGE
            centerPoint = goal.position[0] + ( goal.width / 2 );
            centerPointVert = goal.position[1] - ( goal.height / 2 );
            for ( i = 1; i < mySelection.length; i++ ) {
                // DYELINES IN DOCUMENT
                currItem = mySelection[i];
                // CURRENT DYELINE
                centerPoint = currItem.position[0] + ( currItem.width / 2 );
                centerPointVert = currItem.position[1] - ( currItem.height / 2 );
                newItem = goal.duplicate();
                newItem.position = Array( ( centerPoint - ( goal.width / 2 ) ), ( centerPointVert + ( goal.height/2)));
                newItem.artworkKnockout = currItem.artworkKnockout;
                newItem.clipping = currItem.clipping;
                newItem.isIsolated = currItem.isIsolated;
                newItem.evenodd = currItem.evenodd;
                if ( currItem.polarity ) {
                    newItem.polarity = currItem.polarity;
                }
                newItem.moveBefore( currItem )
                currItem.remove()
            }
            goal.remove()
            // ******** SAVE ******** //
            saveOpts = new PDFSaveOptions();
            saveOpts.compatibility = PDFCompatibility.ACROBAT5; 
            saveOpts.generateThumbnails = true; 
            saveOpts.preserveEditability = true;
            var saveFile = File( sav );
            app.activeDocument.saveAs( saveFile, saveOpts );
            app.activeDocument.close();
            }
    } else {
        alert( 'zemanaDoer() = ELSE' );
    }
}

getOneups( oneupsFolder );
// END SCRIPT