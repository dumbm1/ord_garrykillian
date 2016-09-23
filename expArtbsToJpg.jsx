function scaleAndExp_00 () {
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

function scaleAndExp_01 () {
  var doc       = app.activeDocument,
      abActive  = doc.artboards[doc.artboards.getActiveArtboardIndex ()],
      artWidth  = abActive.artboardRect[2] - abActive.artboardRect[0],
      artHeight = abActive.artboardRect[1] - abActive.artboardRect[3],
      RES_FACT  = 5184, // 72 points in inch
      LONG_SIDE = 5000,
      outRes,
      shortSide;
  if (artWidth >= artHeight) {
    shortSide = (LEN * artHeight) / artWidth;
  } else {
    shortSide = (LEN * artWidth) / artHeight;
  }
  outRes = (LONG_SIDE * shortSide ) / RES_FACT;
  return outRes;
}