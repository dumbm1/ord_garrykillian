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

function scaleAndExp_origin () {
  var doc      = app.activeDocument;
  var abActive = doc.artboards[doc.artboards.getActiveArtboardIndex ()];

  var artWidth  = abActive.artboardRect[2] - abActive.artboardRect[0];
  var artHeight = abActive.artboardRect[1] - abActive.artboardRect[3];

  if (artWidth >= artHeight) {
    var fileName      = doc.fullName.toString ();
    var exportOptions = new ExportOptionsJPEG ();

    var fileSpec                   = new File (fileName);
    exportOptions.antiAliasing     = true;
    exportOptions.artBoardClipping = true;
    exportOptions.qualitySetting   = 100;
    exportOptions.horizontalScale  = (5000 / artWidth) * 100;
    exportOptions.verticalScale    = (5000 / artWidth) * 100;

    doc.exportFile (fileSpec, ExportType.JPEG, exportOptions);
  }
  else {
    var fileName      = doc.fullName.toString ();
    var exportOptions = new ExportOptionsJPEG ();

    var fileSpec                   = new File (fileName);
    exportOptions.antiAliasing     = true;
    exportOptions.artBoardClipping = true;
    exportOptions.qualitySetting   = 100;
    exportOptions.verticalScale    = (5000 / artHeight) * 100;
    exportOptions.horizontalScale  = (5000 / artHeight) * 100;

    doc.exportFile (fileSpec, ExportType.JPEG, exportOptions);
  }
}
