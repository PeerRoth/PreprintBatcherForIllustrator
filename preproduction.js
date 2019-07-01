// PREPRINT BATCHING SCRIPT FOR ADOBE ILLUSTRATOR

// ******  REQUIRED FOLDER STRUCTURE ::
//            C/USER/P/DESK/OJ/Proproduction/_Instructions (ROOTFOLDER)__
//                                                                       |
//      (blank - only folder needed)                                     /Processed    
//      ('Label Example Start USALB.pdf', 'Film Example Start.pdf')      /Starts        
//      (blank - only folder needed)                                     /SavedImprints 
//      ('LABEL.pdf')                                                    /Proofs     
//      ('Dieline_USALB.pdf')                                            /Templates   

// **-- START FILE SHOULD HAVE 'START' AND PRODUCT TYPE ( E.G.- 'file', 'laser'...) IN
// NAME, AND IF 'LABEL' PRODUCT TYPE, THEN ALSO THE TEMPLATE TO BE USED FOR IMPOSITION
// ___________________________________________________________________________________
// FILES & FOLDERS
var rootString          = 'C:\\Users\\peter\\Desktop\\OJ\\Preproduction\\_Instructions';
var processedFolder     = Folder( rootString + '\\Processed' );
var startsFolder        = Folder( rootString + '\\Starts' );
var savedFolder         = Folder( rootString + '\\SavedImprints' );
var proofFolder         = Folder( rootString + '\\Proofs' );
var templatesFolder     = Folder( rootString + '\\Templates');
// ARRAYS
var allPDFFiles         = [ ];
var pdfsWithReg         = [ ];
var regexpMatchArray    = [ [ ] ];  // 2-DIMENSIONAL ARRAY
var productTypeArray    = [ 'film' , 'laser' , 'label' , 'flatbed' ];
var prodTemplatesArray = // DIELINE                                       ROTATION
                    [    //   \/                                             \/
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
                    //  [ [ 'hsr10' ]                                       , 0 ] ,
                        [ [ 'sss8' ]                                        , 0 ] ,
                        [ [ 'es2' ]                                         , 0 ] ,
                        [ [ 'hsc1' ]                                        ,90 ] ,
                        [ [ '2ep' , 'dubcup' , 'hs1z' ]                     , 0 ] ,
                        [ [ '2in' , 'hssk' , 'lb450' , 'lighter' ]          , 0 ] ,
                        [ [ 'ed8' , 'lmroll' , 'mft' , 'tr8901' ]           , 0 ]
                    ];

// MAKE REGEX MATCH [ARRAY]  ( 2-DIMENSIONAL )
for ( var prod = 0 ; prod < productTypeArray.length ; prod ++ ) {
    var doo = new RegExp ( productTypeArray[ prod ], 'i' );
    regexpMatchArray [ prod ] =  [ doo , productTypeArray [ prod ] ]; }

// INTEROGATE FOLDER AND ADD PDFS THAT MATCH 
// REGEX ARRAY TO 'PDFSWITHREGEX' [ARRAY]
function getStartFiles ( startFolder ) {
    var folderContents;
    folderContents = startFolder.getFiles( );
    for ( var y = 0; y < folderContents.length; y ++ ) {
        if ( folderContents[ y ].toLocaleString( ) == "[object Folder]" ) {
            getStartFiles ( Folder( folderContents [ y ].path + "//" + folderContents [ y ].name ) ); }
            if ( folderContents [ y ].name.match( /\.pdf+/ ) && folderContents[ y ].name.match( /Start/i ) )  {
                for ( var reg = 0 ; reg < regexpMatchArray.length ; reg ++ ) {
                        if ( folderContents [ y ].name.match ( regexpMatchArray[ reg ][ 0 ] ) ) {
                            // alert( 'Grabbing PDF: ' + folderContents[ y ].name + ' because it has "Start" and product type in name.' );
                            pdfsWithReg.push( [ folderContents[ y ] , regexpMatchArray[ reg ][ 1 ].toUpperCase( ) ] );
                        } } } }
    iteratePDFs( ) }

function iteratePDFs( ) {
    // alert( 'howd we get here so soonish III:  ' + pdfsWithReg );
    for ( var g = 0; g < pdfsWithReg.length; g ++ ) {
        var goosie = pdfsWithReg[ g ][ 0 ];
        alert( 'GGG:  ' + g + '   .  ' + goosie );
        var goosieSimp = pdfsWithReg[ g ][ 1 ];
        var goosieString = goosie.name;
        // PASS A FLAG THAT WE HAVE TYPE = 'LABEL', SO WE CAN SEND DOWN ADDITIONAL PATH.
        // 'GET LAYERS AND MAKE FILE' FUNCTION NEEDS TO BE CALLED LAST, B/C CLOSES DOCUMENT
        if ( goosieString.match ( /proof/i ) ) {
            alert( 'Not operating on Proof:  ' + goosieSimp + ' -named- ' + goosieString );
        } else {
            if ( goosieSimp.match ( /label/i ) ) {
                makeProof ( goosie , goosieSimp );
                getTemplate ( goosie );
                if ( app.documents.length > 0 ) {
                    app.activeDocument.close( SaveOptions.DONOTSAVECHANGES );
                }  
             } else if ( goosieSimp.match ( /film/i ) ) {
                // alert( 'REFLKTOR cumn' );
                app.open ( goosie );
                if ( goosieString.match ( /reflect/i ) ) {
                    reflectrr( );
                }
                getLayersAndMakeFile ( goosie , goosieSimp );
                if ( app.documents.length > 0 ) {
                    app.activeDocument.close( SaveOptions.DONOTSAVECHANGES );
                }
            } else {
                app.open ( goosie );
                getLayersAndMakeFile ( goosie , goosieSimp ); 
                if ( app.documents.length > 0 ) {
                    app.activeDocument.close( SaveOptions.DONOTSAVECHANGES );
                } } }
        // app.quit( );
    } }

function makeProof( top_file , matcher ) {
    alert( 'In makeProof(). top_file: ' + top_file + '. matcher:   ' + matcher + '  ' + typeof matcher );
        var proofBackground = File( proofFolder + '\\' + matcher + '.pdf' );
        alert( 'proofBackgourn:  ' + proofBackground );
        app.open( proofBackground );

        var activeDocu = app.activeDocument;
        var goalProof = activeDocu.groupItems[ 0 ];
        // var newLayer = activeDocu.layers.add();
        // newLayer.name = 'PROD Layer';
        var uig = activeDocu.placedItems.add();
        uig.file = top_file;
        alert( 'glPos:  ' + goalProof.position[0] + ' . ' + goalProof.position[1]);


    var shaver                          = top_file.toString( );
    var fold_to_shave_string            = 'Starts';
    var fold_to_shave_string_length     = fold_to_shave_string.length;
    var shaveAmount                     = shaver.indexOf( fold_to_shave_string );
    var saveGetFileName                 = shaver.slice( shaveAmount + fold_to_shave_string_length );
    var saveRemoveExtension             = saveGetFileName.replace( '.pdf', '' );
    var saveSuffices                    = '_PROOF';
    var savePattt                        = File( savedFolder + saveRemoveExtension + saveSuffices + '.pdf' );
    // var proofSaved                      = documents.add();


    // var savePattt = File( savedFolder + '\\hihosilva_PROOFLABEL.pdf' );
    var saveOptsss = new PDFSaveOptions();
    saveOptsss.compatibility = PDFCompatibility.ACROBAT5; 
    saveOptsss.generateThumbnails = true; 
    saveOptsss.preserveEditability = true;
    // alert( ' gonnaTrySaveAs= ' + savePattt );
    activeDocu.saveAs( savePattt , saveOptsss );
    activeDocu.close( SaveOptions.DONOTSAVECHANGES );
}

// GRAB 'IMPRINT' LAYER FROM START FILE 
// AND PUT IN NEW DOC, SAVE.
// **LAYBACK IS UNDEFINED UNLESS WE'RE WORKING ON A 'LABEL', WHERE IT = 1
function getLayersAndMakeFile( loose_goose, scooby_doos, layback ) {
    alert( 'Making "NEWIMPRINT" file: ' + typeof loose_goose + ' . ' + loose_goose + ' * ' + scooby_doos + ' * ' + layback );
    layerUnlocker( );
    var lagodact = app.activeDocument;
    var lagolaya = lagodact.layers;
    var targetDoc = documents.add( );
    for ( var l = 0; l < lagolaya.length; l ++ ) {
        var layToMatch = lagolaya[ l ];
        var layToMatchString = layToMatch.name;
        if ( layToMatchString.match( 'IMPRINT' ) ) {
            lagodact.selection = null;
            layToMatch.selected = true;
            var delLayOne = targetDoc.layers.getByName( 'Layer 1' );
            delLayOne.remove( );
            var addLay = targetDoc.layers.add( );
            addLay.name = 'NEW_IMPRINT';
            for ( var b = layToMatch.pageItems.length - 1; b >= 0; b -- ) {
                layToMatch.pageItems[ b ].duplicate( addLay, ElementPlacement.PLACEATBEGINNING );
            } } }
    var prodTypeString = scooby_doos;
    var shaver = loose_goose.toString( );
    var fold_to_shave_string = 'Starts';
    var fold_string_length = fold_to_shave_string.length;
    var shaveAmount = shaver.indexOf( fold_to_shave_string );
    var saveGetFileName = shaver.slice( shaveAmount + fold_string_length );
    var saveRemoveExtension = saveGetFileName.replace( '.pdf', '' );
    var saveSuffices = '_' + prodTypeString + '_' + 'LABEL';
    var savePath = File( savedFolder + '\\' + saveRemoveExtension + saveSuffices + '.pdf' );
    saveOpts = new PDFSaveOptions();
    saveOpts.compatibility = PDFCompatibility.ACROBAT5; 
    saveOpts.generateThumbnails = true; 
    saveOpts.preserveEditability = true;
    alert( ' gonnaTrySaveAs= ' + savePath );
    targetDoc.saveAs( savePath, saveOpts );
    targetDoc.close( SaveOptions.DONOTSAVECHANGES );
    // lagodact.close( SaveOptions.DONOTSAVECHANGES );
}

// fileToReflect
function reflectrr ( ) {
    alert( 'REFLKTOOORRR' );
    layerUnlocker( );
    var a = app.activeDocument;
    var totalMatrix = app.getScaleMatrix( -100 , 100 );
    var e = a.activeLayer.pathItems;
    for ( var pi = 0 ; pi < e.length ; pi ++ ) {
        e[ pi ].transform( totalMatrix );
    } }

function getTemplate( fileToMatch ) {
    alert( 'In getTemplates from FIRST PROJECT, fileToMatch ======= ' + fileToMatch + ' ' + typeof fileToMatch );;
    for ( var i = 0; i < prodTemplatesArray.length; i ++ ) {
        var ouSubarray = prodTemplatesArray[ i ][ 0 ];
        for ( var j = 0; j < ouSubarray.length; j ++ ) {
            var reggie = new RegExp( ouSubarray[ j ], 'i' );
            // alert( 'will we git it  ? reg= ' + reggie + ' . and ftoMatch.NAME: ' + fileToMatch.name );
            if ( fileToMatch.name.match( reggie ) ) {
                var degs = prodTemplatesArray[ i ][ 1 ];
                var templateGot = File( templatesFolder + '\\Dieline_' + ouSubarray[ 0 ].toUpperCase() + '.pdf' );
                alert( 'cooL, we in hea:  ' + degs + ' . ' + templateGot );
                placerDoer( templateGot, fileToMatch, degs );
            } } } }

function placerDoer( tempFi, fileToPlace, degs ) {
    app.open( tempFi );
    layerUnlocker();  
    var oneupsFolderString = rootString + '\\Starts';
    var n = oneupsFolderString.indexOf( '\\Starts ' );
    var fTPSubstring = fileToPlace.name.substring( n - 5 ).replace( '.pdf', '' );
    alert( 'FTP AGeeen??   ' + fTPSubstring );
    var savedFileString = savedFolder + '\\IMPOSED_' + fTPSubstring + '.pdf';
    var processedOneup = File( fileToPlace );
    processedOneup.copy( processedFolder + '\\' + fTPSubstring + '_PROCSD.pdf' );
    var oneupFileForGroup = File( fileToPlace );
    zamena ( savedFileString, oneupFileForGroup, degs );
    // processedOneup.remove( );
}

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
            goal = mySelection[ 0 ]; 
            var degs = degs;
            goal.rotate( degs );
            centerPoint = goal.position[ 0 ] + ( goal.width / 2 );
            centerPointVert = goal.position[1] - ( goal.height / 2 );
            for ( i = 1; i < mySelection.length; i++ ) {
                currItem = mySelection[ i ];
                centerPoint = currItem.position[ 0 ] + ( currItem.width / 2 );
                centerPointVert = currItem.position[ 1 ] - ( currItem.height / 2 );
                newItem = goal.duplicate( );
                newItem.position = Array( ( centerPoint - ( goal.width / 2 ) ), ( centerPointVert + ( goal.height/2)));
                newItem.artworkKnockout = currItem.artworkKnockout;
                newItem.clipping = currItem.clipping;
                newItem.isIsolated = currItem.isIsolated;
                newItem.evenodd = currItem.evenodd;
                if ( currItem.polarity ) {
                    newItem.polarity = currItem.polarity; }
                newItem.moveBefore( currItem )
                currItem.remove( ) }
            goal.remove()
            alert( 'zam, right after removegoal' );
            saveOpts = new PDFSaveOptions();
            saveOpts.compatibility = PDFCompatibility.ACROBAT5; 
            saveOpts.generateThumbnails = true; 
            saveOpts.preserveEditability = true;
            var saveFile = File( sav );
            alert( 'SAVING :  ' + saveFile + ' , w/ options:  ' + saveOpts );
            app.activeDocument.saveAs( saveFile, saveOpts );
            // app.activeDocument.close();
            }
    } else {
        alert( 'zemanaDoer() = ELSE' );
    } }

// UNLOCK LAYERS, CALLED FROM GETLAYERSMAKEFILE()
function layerUnlocker( ) {
    if ( app.documents.length > 0 ) {
        // alert( 'unlocking lays' );
        var actDoc = app.activeDocument;
        var layersInDoc = actDoc.layers;
        for ( i = 0; i < layersInDoc.length; i ++ ) {
            layersInDoc[ i ].locked = false;
        } } }

// SPECIFY FOLDER TO INTEROGATE, START SCRIPT
getStartFiles( startsFolder );