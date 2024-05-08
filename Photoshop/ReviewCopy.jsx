/**
 * Review Copy
 *
 * @author        Johnny Storm
 * @version        2.2.2
 *
 * 2021-11-24: Initial Version.
 *
 **/

/*
Licensed under the MIT License

Copyright (c) 2022 John D. Roberts

Permission is hereby granted, free of charge, to any person obtaining a copy of
this software and associated documentation files (the "Software"), to deal in
the Software without restriction, including without limitation the rights to
use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of
the Software, and to permit persons to whom the Software is furnished to do so,
subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS
FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR
COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER
IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN
CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

*/

/*
@@@BUILDINFO@@@ Review Copy.jsx 2.2.2.0
*/

// enable double clicking from the Macintosh Finder or the Windows Explorer
#target photoshop

Object.prototype.writePlist = function () {
    var dict = new XML('<dict/>');
    for (var key in this) {
        if (typeof this[key] != 'function') {
            var node = new XML('<key>'+key+'</key>');
            dict.appendChild(node);
            var string = new XML('<string>'+this[key]+'</string>');
            dict.appendChild(string);
        }
    }

    var root = new XML('<plist/>');
        root.@version = '1.0';
        root.appendChild(dict);

    var plist = new File(Folder.userData + '/' + PREFS);
        plist.open('w');
        plist.writeln(root.toXMLString());
        plist.close();
};

const CompressionType = {
    JPEG     : 12,
    ZIP      : 65540,
    JPEG2000 : 14,
    MINIMUM  : 11,
    LOW      : 10,
    MEDIUM   : 9,
    HIGH     : 8,
    MAXIMUM  : 7,
    NONE     : 1,

};

const MAX_WIDTH        = 1988;
const MAX_HEIGHT       = 3056;
const VERSION          = '2.2.2';
const WATERMARK        = 'REVIEW COPY';
const RESOLUTION       = 72;
const REVIEW_HEIGHT    = 2000;
const FONTS            = ['Futura', 'Helvetica', 'Arial'];

var crop, x, y, width, height, selectedFile, selectedFolder;
var pageHeight        = 0;
var crop              = 'cropBox';
var trim              = 'none';
var resolution        = 300;
var pageCount         = 6000;
var resize            = true;
var portrait          = false;
var fitto             = false;
var left              = 0;
var top               = 0;
var right             = 0;
var bottom            = 0;
var compression       = 70;
var startPage         = 1;

var exportPDF         = true;
var reviewCopy        = false;
var exportTIFF        = false;
var exportJPEG        = false;

const OS             = $.os.toLowerCase().indexOf('mac') >= 0 ? "MAC": "WINDOWS";
const PREFS          = 'com.johnnystorm.reviewCopy.plist';
const DEFAULT_FOLDER = Folder.desktop;

var prefs            = readPlist();

var defaultPath      = prefs.defaultPath ? prefs.defaultPath : DEFAULT_FOLDER;
var watermark        = prefs.watermark   ? prefs.watermark   : WATERMARK;
var defaultFont      = findFonts();

var originalRulerUnits                    = app.preferences.rulerUnits;
app.preferences.rulerUnits                = Units.PIXELS;
app.displayDialogs                        = DialogModes.ERROR;

const exportJPEGOptions                   = new ExportOptionsSaveForWeb();
    exportJPEGOptions.format              = SaveDocumentType.JPEG;
    exportJPEGOptions.includeProfile      = false;
    exportJPEGOptions.optimized           = true;
    exportJPEGOptions.quality             = compression;

const tiffSaveOptions = new TiffSaveOptions();
    tiffSaveOptions.embedColorProfile = true;
    tiffSaveOptions.alphaChannels = false;
    tiffSaveOptions.layers = false;
    tiffSaveOptions.imageCompression = TIFFEncoding.TIFFLZW;

const pdfSaveOptions = new PDFSaveOptions();
    pdfSaveOptions.embedColorProfile = true;
    pdfSaveOptions.encoding = PDFEncoding.PDFZIP;

var dlg = createDlg();

app.preferences.rulerUnits              = originalRulerUnits;

function createDlg() {
    var selectedFile;
    var dlg                             = new Window('dialog', 'ReviewCopies');

    this.body                           = dlg.add("panel {orientation: 'column'}");
    this.body.alignment                 = 'fill'

    this.row1                           = this.body.add('group');
    this.row1.orientation               = 'row';
    this.row1.alignment                 = 'fill'

    this.left                           = this.row1.add('group');
    this.left.orientation               = 'column';
    this.left.alignment                 = 'fill';

    this.settings                       = this.left.add('panel', undefined, 'Settings');
    this.settings.minimumSize           = [290, 0];
    this.settings.orientation           = 'column';
    this.settings.alignChildren         = ['left', 'top'];

    this.dimension                      = this.settings.add("group {orientation: 'row'}");
    this.dimension.s                    = this.dimension.add("StaticText { text:'Crop To:', minimumSize: [75,0], alignment: 'right' }");
    this.crop                           = this.dimension.add("DropDownList", undefined, [ 'Trim Box', 'Crop Box', 'Bounding Box', 'Media Box' ]);
    this.crop.items[ 0 ].selected       = true;

    this.trimBtn                        = this.settings.add("Group {orientation: 'row', alignment: 'left'}");
    this.trimBtn.s                        = this.trimBtn.add("StaticText { text:'Trim To:', minimumSize: [75,0], alignment: 'right' }");
    this.options                        = this.trimBtn.add("DropDownList", undefined, [ 'None', 'Transparent Pixels', 'Top Left Pixel Color', 'Bottom Right Pixel Color' ]);
    this.options.items[ 0 ].selected    = true;

    this.resolution                     = this.settings.add("Group {orientation: 'row', alignment: 'left'}");
    this.resolution.s                   = this.resolution.add("StaticText { text:'Resolution:', minimumSize: [75,0] }");
    this.resolution.e                   = this.resolution.add("EditText { characters: 7, text: 300 }");

    this.compression                       = this.settings.add("Group {orientation: 'row', alignment: 'left'}");
    this.compression.s                     = this.compression.add("StaticText { text:'Quality:', minimumSize: [75,0] }");
    this.compression.e                     = this.compression.add("DropDownList", undefined, [ '50 - Low', '60 - Medium', '70 - Standard', '80 - High', '90 - Very High', '100 - Maximum' ]);
    this.compression.e.items[ 2 ].selected = true;

    this.start                             = this.settings.add("Group {orientation: 'row', alignment: 'left'}");
    this.start.s                           = this.start.add("StaticText { text:'Start:', minimumSize: [75,0] }");
    this.start.e                           = this.start.add("EditText { characters: 7, text: " + startPage + " }");

    this.watermark                         = this.settings.add("Group {orientation: 'row', alignment: 'left'}");
    this.watermark.s                       = this.watermark.add("StaticText { text:'Watermark:', minimumSize: [75,0] }");
    this.watermark.e                       = this.watermark.add("EditText { characters: 20, text: '" + WATERMARK + "' }");
    this.watermark.e.alignment             = 'fill';

    this.buttons                           = this.left.add('group');
    this.buttons.orientation               = 'row';

    this.selectButton                = this.buttons.add('button', undefined, 'Open File...', {name:'open'});
    this.selectButton.align          = 'left';
    this.selectButton.onClick        = function() {
        var path = new File(defaultPath);
        selectedFile = path.openDlg("Choose File:","*.pdf");
        while (selectedFile.alias) {
            selectedFile = file.resolve().openDlg("Choose File:");
        }
        uri.text = File.decode(selectedFile.name);
        prefs.defaultPath = selectedFile.path;
        prefs.writePlist();
    };

    this.selectFolder                    = this.buttons.add('button', undefined, 'Open Folder...', {name:'openFolder'});
    this.selectFolder.align             = 'right';
    this.selectFolder.onClick            = function() {
        var folder = new Folder();
        if(selectedFolder = folder.selectDlg(undefined, defaultPath)) {
            uri.text = File.decode(selectedFolder.name);
            prefs.defaultPath = selectedFolder.path;
            prefs.writePlist();
        }
    };

    this.right                            = this.row1.add('group');
    this.right.orientation                = 'column';
    this.right.alignment                = 'top';

    this.naming                            = this.right.add('panel', undefined, 'Naming Convention');
    this.naming.orientation                = 'column';
    this.naming.alignment                = 'top';
    this.naming.alignChildren            = 'left';
    this.naming.minimumSize                = [175, 0];
    this.PDFFilename                    = naming.add('checkbox', undefined, ' PDF Filename');
    this.PDFFilename.value                = true;
    this.prefix                            = this.naming.add('checkbox', undefined, ' Ordering Prefix');
    this.prefix.value                    = false;

    this.presets                        = this.right.add('panel', undefined, 'Export');
    this.presets.orientation            = 'column';
    this.presets.alignment                = 'top';
    this.presets.alignChildren            = 'left';
    this.presets.minimumSize            = [175, 0];

    this.exportPDF                      = presets.add('checkbox', undefined, ' PDF');
    this.exportPDF.value                = true;
    this.exportJPEG                     = presets.add('checkbox', undefined, ' JPEGs');
    this.exportTIFF                     = presets.add('checkbox', undefined, ' TIFFs');
    this.review                         = presets.add('checkbox', undefined, ' Review Copy');

    this.row2                           = this.body.add('group');
    this.row2.orientation               = 'stack';
    this.row2.alignment                 = 'fill'

    this.uri                           = this.row2.add('statictext', undefined, 'No file chosen', {readonly:true,multiline:false});
    this.uri.alignment                 = 'fill';

    this.footer                         = dlg.add('group');
    this.footer.orientation             = 'row';
    this.footer.alignment               = 'fill';

    this.leftBtn                        = footer.add('group');
    this.leftBtn.orientation            = 'row';
    this.leftBtn.alignChildren          = 'left';
    this.leftBtn.minimumSize            = [300, 0];

    this.rightBtn                       = this.footer.add('group');
    this.rightBtn.orientation           = 'row';
    this.rightBtn.alignChildren         = 'right';
    this.rightBtn.alignment             = ['fill', 'right'];

    this.version              = this.leftBtn.add('statictext', undefined, 'Version: ' + VERSION, {readonly:true,multiline:false});
    this.cancelBtn            = this.rightBtn.add('button', undefined, 'Cancel', {name:'cancel'});
    this.okBtn                = this.rightBtn.add('button', undefined, 'OK', {name:'ok'});

    dlg.center();
    if ( 1 == dlg.show() ) {
        prefs.watermark = this.watermark.e.text;
        prefs.writePlist();
        if (!this.exportPDF.value && !this.exportJPEG.value && !this.exportTIFF.value && !this.review.value) {
            alert('Please select an export option.');
        } else {
            try {
                if(selectedFile != null) {
                    defineSettings(this);
                    if(processFile(this, selectedFile)) {
                        alert('Your file is complete.');
                    }
                } else if(selectedFolder != null) {
                    defineSettings(this);
                    processFolder(this, selectedFolder);
                } else {
                    alert("Please Select a File to Process");
                }
            } catch(e) {
                alert(e);
            }
        }
    }
    return dlg;
}

function defineSettings(dlg) {
    switch(dlg.crop.selection.text) {
        case 'Trim Box' :
            crop    = 'trimBox';
            break;
        case 'Bounding Box' :
            crop    = 'boundingBox';
            break;
        case 'Media Box' :
            crop    = 'mediaBox';
            break;
        default :
            crop    = 'cropBox';
    }

    switch(dlg.compression.e.selection.index) {
        case 0 :
            compression = 50;
            break;
        case 1 :
            compression = 60;
            break;
        case 2 :
            compression = 70;
            break;
        case 3 :
            compression = 80;
            break;
        case 4 :
            compression = 90;
            break;
        case 5 :
            compression = 100;
            break;
        default :
            compression = 70;
    }

    resolution  = dlg.resolution.e.text;
    startPage   = dlg.start.e.text;
    reviewCopy  = dlg.review.value;
    trim        = dlg.options.selection.text;
    watermark   = dlg.watermark.e.text;
    exportPDF   = dlg.exportPDF.value;
    exportTIFF  = dlg.exportTIFF.value;
    exportJPEG  = dlg.exportJPEG.value;
}

function processFolder(dlg, folder) {
    var files = selectedFolder.getFiles("*.pdf");
        files.sort();

    for (var i = 0; i < files.length; i++) {
        var file = files[i];
        try {
            processFile(dlg, file);
        } catch(e) {
            alert(e);
        }

    }
    alert('Your files are complete.');
}

function processFile(dlg, file) {

    var comicFull    = file.name;

    var temp = comicFull.split('.');
        temp.pop();

    var comicName = temp.join('.');
        comicName = comicName.replace(' ', '-');

    var directory = file.path;

    var parentFolder = new Folder(directory + '/' + comicName);
        parentFolder.create();
    var res = getResolution(file, 1);

    if (res > resolution) res = resolution;


    if(parentFolder) {
        var pages                = [];
        var fileList = [];

        var count = pageCount > 0 ? pageCount : 6000;

        for(var n = startPage; n <= count; n++) {

            var action = rasterizePDF(n, res, file, crop);
            if(app.documents.length == 0 || action.count == 0) break;

            if(n < 10) {
                z = '00' + n;
            } else if(n < 100) {
                z = '0' + n;
            } else {
                z = n;
            }

            var filename    = (dlg.prefix.value == true) ? "2-"+comicName+'-'+z+".jpg" : comicName+'-'+z+".jpg";
            var jpegFile    = dlg.PDFFilename.value ? new File(parentFolder.fullName+'/'+filename) : new File(parentFolder.fullName+'/'+z+'.jpg');
            var tiffFile    = new File(jpegFile.path + '/' + getFilename(jpegFile.name) + '.tif');

            fileList.push(tiffFile);

            processImage(jpegFile, tiffFile);
        }

        if (exportPDF) {
            var pdfFilename = reviewCopy ? comicName + '-review-'+randomExt()+'.pdf' : comicName + '-digital.pdf';
            var digitalFile = new File(parentFolder.parent.fullName +'/' + pdfFilename);

            exportPDFFile(fileList, digitalFile);
        }
        if (!exportTIFF) {
            deleteFiles(parentFolder);
        }
        if (!exportTIFF && !exportJPEG) {
            parentFolder.remove();
        }
        return true;
    } else {
        return false;
        alert('Directory ' + directory + '/' + comicName + " doesn't exist");
    }
}

function randomExt() {
    var list = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789';

    var a = list.charAt((Math.random()*list.length)-1);
    var b = list.charAt((Math.random()*list.length)-1);
    var c = list.charAt((Math.random()*list.length)-1);
    return a + b + c;
}

function processImage(jpegFile, tiffFile) {
    app.preferences.rulerUnits = Units.PIXELS
    var doc         = app.activeDocument;

    var width        = Math.floor(doc.width.value);
    var height        = Math.floor(doc.height.value);
    switch(trim) {
        case 'Transparent Pixels' :
            doc.trim(TrimType.TRANSPARENT,true,true,true,true);
            break;
        case 'Top Left Pixel Color' :
            doc.trim(TrimType.TOPLEFT, true, false, true, false);
            break;
        case 'Bottom Right Pixel Color' :
            doc.trim(TrimType.BOTTOMRIGHT, true, false, true, false);
            break;
    }
    if ((height > MAX_HEIGHT) === true) {
        cropImage();
    }

    if (reviewCopy) {
        resizeImage();
        addWatermark();
    } else {
        app.activeDocument.resizeImage(doc.width.value, doc.height.value, RESOLUTION);
    }

    doc.saveAs(tiffFile, tiffSaveOptions, true, Extension.LOWERCASE);

    if (exportJPEG) {
        exportJPEGOptions.quality = compression;
        app.activeDocument.exportDocument(jpegFile, ExportType.SAVEFORWEB, exportJPEGOptions);
    }
    app.activeDocument.close(SaveOptions.DONOTSAVECHANGES);
}

function cropImage() {
    var width       = app.activeDocument.width.value;
    var height      = app.activeDocument.height.value;
    var the_width   = width > MAX_WIDTH * 1.5 ? MAX_WIDTH * 2 : MAX_WIDTH;
    var crop_x      = parseInt((width-the_width)*.5);
    var crop_y      = parseInt((height-MAX_HEIGHT)*.5);
    var crop_width  = the_width + crop_x;
    var crop_height = MAX_HEIGHT + crop_y;
    var crop_to     = new Array(crop_x, crop_y, crop_width, crop_height);

    app.activeDocument.crop(crop_to, 0, the_width, MAX_HEIGHT);
}

function resizeImage() {
    var width        = app.activeDocument.width.value;
    var height        = app.activeDocument.height.value;
    var the_width    = width > MAX_WIDTH * 1.5 ? 1301 * 2 : 1301;

    app.activeDocument.resizeImage(the_width, REVIEW_HEIGHT, RESOLUTION);
}

function addURL(position) {
    var textColor = new SolidColor();
        textColor.rgb.red   = 255;
        textColor.rgb.green = 255;
        textColor.rgb.blue  = 255;

    var layer = app.activeDocument.artLayers.add();
        layer.opacity                 = 50;
        layer.kind                    = LayerKind.TEXT;
        layer.textItem.font           = defaultFont;
        layer.textItem.useAutoLeading = false;
        layer.textItem.size           = 24;
        layer.textItem.leading        = 24;
        layer.textItem.color          = textColor;
        layer.textItem.justification  = Justification.CENTER;
        layer.textItem.contents       = watermark;
        layer.textItem.position       = position;
        layer.blendMode               = BlendMode.DIFFERENCE;

    return layer;
}

function addWatermark() {
    var width        = app.activeDocument.width.value;
    var height        = app.activeDocument.height.value;

    var header = addURL([width * .5,36])
    var footer = addURL([width * .5,height - 24])

    app.activeDocument.flatten();
}

function rasterizePDF(pageNumber, res, pdfFile, crop) {
    var cropto    = crop ? crop : 'cropBox'; //trimBox
    var mode    = 'RGBC';//OpenDocumentMode.RGB;

    var desc = new ActionDescriptor();

    var optionsDesc = new ActionDescriptor();
        optionsDesc.putString( charIDToTypeID( "Nm  " ), "rasterized page" );
        optionsDesc.putEnumerated( charIDToTypeID( "Crop" ), stringIDToTypeID( "cropTo" ), stringIDToTypeID( cropto ) );
        if(res > 0) optionsDesc.putUnitDouble( charIDToTypeID( "Rslt" ), charIDToTypeID( "#Rsl" ), res);
        optionsDesc.putEnumerated( charIDToTypeID( "Md  " ), charIDToTypeID( "ClrS" ), charIDToTypeID( mode ) );
        optionsDesc.putBoolean( charIDToTypeID( "AntA" ), true );
        optionsDesc.putBoolean( stringIDToTypeID( "suppressWarnings" ), false );
        optionsDesc.putEnumerated( charIDToTypeID( "fsel" ), stringIDToTypeID( "pdfSelection" ), stringIDToTypeID( "page"  ));
        optionsDesc.putInteger( charIDToTypeID( "PgNm" ), pageNumber );
    desc.putObject( charIDToTypeID( "As  " ), charIDToTypeID( "PDFG" ), optionsDesc );
    desc.putPath( charIDToTypeID( "null" ), pdfFile );

    return executeAction( charIDToTypeID( "Opn " ), desc, DialogModes.NO );
}

function exportPDFFile(fileList, file) {
    // run PDF Presentation
    var presentationOptions = new PresentationOptions();
        presentationOptions.presentation = false;
        presentationOptions.view = true;
        presentationOptions.PDFFileOptions = pdfSaveOptions;
        presentationOptions.includeFilename = false;
    app.makePDFPresentation(fileList, file, presentationOptions);
}

function deleteFiles(dir) {
    var files = dir.getFiles("*.tif");
    for (var i = 0; i < files.length; i++) {
        var file = files[i];
        if (file.exists) {
            file.remove();
        }
    }
}

function getResolution(fileRef, page) {
    var pdfOpenOptions                  = new PDFOpenOptions();
        pdfOpenOptions.page             = page;
        pdfOpenOptions.suppressWarnings = true;
        pdfOpenOptions.usePageNumber    = false;

    var doc = app.open( fileRef, pdfOpenOptions );

    var resolution = doc.resolution;
    doc.close(SaveOptions.DONOTSAVECHANGES);
    return Math.round(resolution);
}

function getFilename(str) {
    return str.substring(str.lastIndexOf("\\") + 1, str.lastIndexOf("."));
};

function findFonts() {
    for (var i = 0; i < FONTS.length; i++ ) {
        var font = findFont( FONTS[i] );
        if ( font ) return font;
    }
}

function findFont( fontName ) {
    var installedFonts = app.fonts;
    for ( var i = 0; i < installedFonts.length; i++ ) {
        if ( installedFonts[i].name == fontName  ) {
            return installedFonts[i];
        }
    }
    return null;
}

function readPlist() {
    var plist = new File(Folder.userData + '/' + PREFS);
        plist.open('r');

    var prefs = plist.read();

        plist.close();

    var doc = new XML(prefs);

    var list = doc.dict.children();
    var key = null;
    var o = {};
    for (var i= 0 ; i < list.length(); i++) {
        if (list[i].name() == 'key') {
            key = list[i];
        } else {
            o[key] = list[i];
        }
    }
    return o;
};
