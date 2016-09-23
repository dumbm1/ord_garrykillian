var doc = app.activeDocument;  
var abActive  = doc.artboards[doc.artboards.getActiveArtboardIndex()];  
  
  
var artWidth = abActive.artboardRect[2] - abActive.artboardRect[0];  
var artHeight = abActive.artboardRect[1] - abActive.artboardRect[3];  
  
  
if (artWidth>=artHeight)  
{  
   var fileName = doc.fullName.toString();  
    var exportOptions = new ExportOptionsJPEG();  
     
    var fileSpec = new File(fileName);  
    exportOptions.antiAliasing = true;  
    exportOptions.artBoardClipping = true;  
    exportOptions.qualitySetting = 100;  
    exportOptions.horizontalScale = (5000/artWidth)*100;  
    exportOptions.verticalScale = (5000/artWidth)*100;  
     
    doc.exportFile( fileSpec, ExportType.JPEG, exportOptions );  
   }  
else  
{  
    var fileName = doc.fullName.toString();  
    var exportOptions = new ExportOptionsJPEG();  
     
    var fileSpec = new File(fileName);  
    exportOptions.antiAliasing = true;  
    exportOptions.artBoardClipping = true;  
    exportOptions.qualitySetting = 100;  
    exportOptions.verticalScale = (5000/artHeight)*100;  
    exportOptions.horizontalScale = (5000/artHeight)*100;  
     
     
    doc.exportFile( fileSpec, ExportType.JPEG, exportOptions );  
}  