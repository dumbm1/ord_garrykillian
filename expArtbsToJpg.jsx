/**
 * Autor: garrykillian
 * Then creating action with this script and batch for multiple files.
 * For batch convert PNG to JPG use XnConvert.
 * */
function expToJpgViaPng () {
  var doc      = app.activeDocument;
  var abActive = doc.artboards[doc.artboards.getActiveArtboardIndex ()];

  var artWidth  = abActive.artboardRect[2] - abActive.artboardRect[0];
  var artHeight = abActive.artboardRect[1] - abActive.artboardRect[3];

  if (artWidth >= artHeight) {
    var fileName      = doc.fullName.toString ();
    var exportOptions = new ExportOptionsPNG24 ();

    var fileSpec                   = new File (fileName);
    exportOptions.antiAliasing     = true;
    exportOptions.artBoardClipping = true;
    exportOptions.transparency     = false;
    exportOptions.qualitySetting   = 100;
    exportOptions.horizontalScale  = (5000 / artWidth) * 100;
    exportOptions.verticalScale    = (5000 / artWidth) * 100;

    doc.exportFile (fileSpec, ExportType.PNG24, exportOptions);
  }
  else {
    var fileName      = doc.fullName.toString ();
    var exportOptions = new ExportOptionsPNG24 ();

    var fileSpec                   = new File (fileName);
    exportOptions.antiAliasing     = true;
    exportOptions.artBoardClipping = true;
    exportOptions.transparency     = false;
    exportOptions.qualitySetting   = 100;
    exportOptions.verticalScale    = (5000 / artHeight) * 100;
    exportOptions.horizontalScale  = (5000 / artHeight) * 100;

    doc.exportFile (fileSpec, ExportType.PNG24, exportOptions);
  }
  doc.close (SaveOptions.DONOTSAVECHANGES);
}

// todo: convert the resolution to 72 dpi in photoshop
function expToJpgViaPdf () {
  var d                  = activeDocument;
  var storeInteractLavel = app.userInteractionLevel;

  var folderPath = new Folder (activeDocument.path);

  (new Folder (folderPath).exists == false) ? new Folder (folderPath).create () : '';

  var fileName = d.name.slice (0, d.name.lastIndexOf ('.')),
      fullPath = folderPath + '/' + fileName,
      artbsLen = d.artboards.length,
      res      = [];

  app.userInteractionLevel = UserInteractionLevel.DONTDISPLAYALERTS;

  for (var j = 0; j < artbsLen; j++) {
    activeDocument.artboards.setActiveArtboardIndex (j);
    res.push (getOpenPhotoshopResolution ());
  }
  saveAsPdf (fullPath);
  makeJpgFromPdf (fullPath, artbsLen, res);

  app.userInteractionLevel = storeInteractLavel;

  function getOpenPhotoshopResolution () {
    var d         = activeDocument,
        artbAct   = d.artboards[d.artboards.getActiveArtboardIndex ()],
        artbW     = artbAct.artboardRect[2] - artbAct.artboardRect[0],
        artbH     = artbAct.artboardRect[1] - artbAct.artboardRect[3],
        LONG_SIDE = 5000,
        BASE_RES  = 72,
        openShopRes;

    if (artbW >= artbH) {
      openShopRes = LONG_SIDE * BASE_RES / artbW;
    } else {
      openShopRes = LONG_SIDE * BASE_RES / artbH;
    }
    return openShopRes;
  }

  function saveAsPdf (fullPath) {
    var pdfSaveOpts = new PDFSaveOptions (),
        f           = new File (fullPath);

    pdfSaveOpts.PDFPreset = '[Illustrator Default]';
    /**
     * COLORCONVERSIONREPURPOSE COLORCONVERSIONTODEST None
     * */
    pdfSaveOpts.colorConversionID = ColorConversion.COLORCONVERSIONREPURPOSE;
    pdfSaveOpts.viewAfterSaving = false;

    d.saveAs (f, pdfSaveOpts);
  }

  function makeJpgFromPdf (fullPath, artbsLen, res) {
//    var btCount = 1;
    sendBt ();

    function sendBt () {
      /* $.writeln (
       '\n' + new Array ( btCount ).join ( '-' ) + '->\n' +
       new Array ( btCount ).join ( '-' ) + '-> ' + fileName + ' sending bt message #' + btCount++ + '\n' +
       new Array ( btCount - 1 ).join ( '-' ) + '->\n'
       );*/
      var bt       = new BridgeTalk ();
      bt.target    = 'photoshop';
      bt.body      = _ps_makeJpg.toString () +
        ';_ps_makeJpg("' + fullPath + '","' + artbsLen + '","' + res.join ('::') + '");';
      bt.timeout   = 1200;
      bt.onTimeout = function () {
        sendBt ();
      }
      return bt.send ();
    }

    function _ps_makeJpg (fullPath, artbsLen, res) {

      var res = res.split ('::');

      app.displayDialogs = DialogModes.NO;

      var pdfFile     = new File (fullPath + '.pdf'),
          jpgPath,
          pdfOpenOpts = new PDFOpenOptions,
          jpgSaveOpts = new JPEGSaveOptions ();

      pdfOpenOpts.usePageNumber = true;

      pdfOpenOpts.antiAlias        = true;
      pdfOpenOpts.bitsPerChannel   = BitsPerChannelType.EIGHT;
      pdfOpenOpts.cropPage         = CropToType.CROPBOX;
      pdfOpenOpts.mode             = OpenDocumentMode.RGB;
      pdfOpenOpts.suppressWarnings = true;

      jpgSaveOpts.embedColorProfile = false;
      jpgSaveOpts.formatOptions     = FormatOptions.STANDARDBASELINE; // OPTIMIZEDBASELINE PROGRESSIVE STANDARDBASELINE
      jpgSaveOpts.matte             = MatteType.NONE // BACKGROUND BLACK FOREGROUND NETSCAPE NONE SEMIGRAY WHITE
      jpgSaveOpts.quality           = 11; // number [0..12]
//      jpgSaveOpts.scans = 3; // number [3..5] only for when formatOptions = FormatOptions.PROGRESSIVE

      try {
        for (var i = 1; i < artbsLen + 2; i++) {
          ( i < 10 ) ? jpgPath = fullPath + '-0' + i : jpgPath = fullPath + '-' + i;
          pdfOpenOpts.page       = i;
          pdfOpenOpts.resolution = res[i];
          app.open (pdfFile, pdfOpenOpts);
          app.activeDocument.saveAs (new File (jpgPath), jpgSaveOpts, true);
          app.activeDocument.close (SaveOptions.DONOTSAVECHANGES);
        }
      } catch (e) {
        alert (e);
      } finally {
        pdfFile.remove ();
      }
    }
  }
}
