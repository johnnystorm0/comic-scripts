/**
 * Merge Pages
 * The Merge Pages script will allow you to select two images and merge them into a single image
 * the width of both pages.
 *
 * @author		Johnny Storm
 * @version		1.0.0
 *
 * 2024-04-18: Initial Version.
 *
 **/

/*
Licensed under the MIT License

Copyright (c) 2024 John D. Roberts

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
@@@BUILDINFO@@@ MergePages.jsx 1.0.0.0
*/
#target photoshop
#script 'Merge Pages';

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

	var plist = new File(PREFS_FOLDER + '/' + PREFS);
		plist.open('w');
		plist.writeln(root.toXMLString());
		plist.close();

};
const VERSION	     = "1.0.0";
const OS             = $.os.toLowerCase().indexOf('mac') >= 0 ? "MAC": "WINDOWS";
const PREFS_FOLDER   = OS == 'MAC' ? '~/Library/Preferences' : '/AppData/Roaming';
const PREFS          = 'com.johnnystorm.default.plist';
const DEFAULT_FOLDER = OS == 'MAC' ? '~/Desktop' : 'C:\Users\Default\Desktop';

var prefs            = readPlist();
var defaultPath	     = prefs.defaultPath ? prefs.defaultPath : DEFAULT_FOLDER;

const fileMask = function(file) {
	return file instanceof Folder || file.name.match(/\.jpg$/i) ? true : false;
}

var originalRulerUnits					= app.preferences.rulerUnits;
app.preferences.rulerUnits				= Units.PIXELS;
app.displayDialogs						= DialogModes.ERROR;

const bounds = [10,40,300,60];

const exportJPEGOptions					= new ExportOptionsSaveForWeb();
	exportJPEGOptions.format			= SaveDocumentType.JPEG;
	exportJPEGOptions.includeProfile	= false;
	exportJPEGOptions.optimized			= true;
	exportJPEGOptions.quality			= 70;

createDlg();
app.preferences.rulerUnits = originalRulerUnits;

function createDlg() {
	var dlg	                        = new Window('dialog', 'Merge Pages');

	var body                        = dlg.add("panel {orientation: 'column'}");

	var header = body.add('group');
		header.orientation            = 'row';
		header.alignChildren	      = 'left';
		header.alignment              = 'fill'

	var row1 = addRow(body, 'Left');
	var row2 = addRow(body, 'Right');

	var buttons						= dlg.add('group');
		buttons.orientation			= 'row';
		buttons.alignment           = 'fit';

        buttons.left				= buttons.add('group');
        buttons.left.orientation	= 'row';
        buttons.left.alignChildren	= 'left';
        buttons.left.minimumSize	= [300, 0];
        buttons.right				= buttons.add('group');
        buttons.right.orientation	= 'row';
        buttons.right.alignChildren	= 'right';

    var options                     = body.add('group');
        options.orientation         = 'row';
        options.alignChildren	    = 'left';
        options.alignment           = 'fill';

    var keepOpen                    = options.add('checkbox', undefined, ' Keep file open after merge.');
        keepOpen.value              = prefs.keepOpen;

	var version		            = buttons.left.add('statictext', undefined, 'Version: ' + VERSION, {readonly:true,multiline:false});
	var cancelBtn		        = buttons.right.add('button', undefined, 'Cancel', {name:'cancel'});
	var okBtn			        = buttons.right.add('button', undefined, 'OK', {name:'ok'});
		okBtn.onClick = function() {
			if (row1.file == null) {
				alert('Please select a file for the left page.');
			} else if (row2.file == null) {
				alert('Please select a file for the right page.');
			} else {
                prefs.keepOpen = keepOpen.value;
				prefs.writePlist();

				okBtn.enabled = false;
				dlg.close();
				try {
					mergeFiles(row1, row2);
				} catch(e) {
					alert(e);
				}
			}
		};

	dlg.center();
	dlg.show()
}

function mergeFiles(row1, row2) {
	var doc = app.open(row1.file);
	var w =  app.activeDocument.width.value;
	var h =  app.activeDocument.height.value;
	doc.resizeCanvas(w*2, h, AnchorPosition.TOPLEFT);
	placeFile(row2.file);
	alignToCanvas(true, "ADSRights");

    if (!prefs.keepOpen) {
        saveFile(row1.file);
    }
}

function saveFile(file) {
	var name	= file.name;
	var temp = name.split('.');
	var ext = temp.pop();
	var filename = temp.join('.') + '-merged.jpg';
	var directory = new Folder(file.path);

	try {
		var newFile = new File(directory.fullName+'/'+filename);
		app.activeDocument.exportDocument(newFile, ExportType.SAVEFORWEB, exportJPEGOptions);
		app.activeDocument.close(SaveOptions.DONOTSAVECHANGES);
	} catch(e) {
		alert(e);
	} finally {
		alert('Done');
	}
}

function addRow(body, title) {
	var row = body.add('group');
		row.orientation            = 'row';
		row.minimumSize.width      = 450;
		row.alignment              = 'fill';

		row.title                  = row.add("StaticText { minimumSize: [50, 0], alignment: 'right' }");
		row.title.text             = title + ':';
		row.uri                    = row.add("EditText { minimumSize: [300,0], enabled:false }");
		row.btn                    = row.add('button', undefined, 'Open File...', {name:'open1'});
		row.btn.onClick	= function() {
			var path = new File(defaultPath);
			row.file = path.openDlg("Choose Files:",fileMask, false);
			if (row.file != null) {
                prefs.defaultPath = row.file.path;
				prefs.writePlist();
				defaultPath = row.file.path;
                alert(defaultPath);
				row.uri.text = File.decode(row.file.path);
			}
		}

	return row;
}

function placeFile(placeFile) {
    var desc21 = new ActionDescriptor();
    desc21.putPath( charIDToTypeID('null'), new File(placeFile) );
    desc21.putEnumerated( charIDToTypeID('FTcs'), charIDToTypeID('QCSt'), charIDToTypeID('Qcsa') );
    var desc22 = new ActionDescriptor();
    desc22.putUnitDouble( charIDToTypeID('Hrzn'), charIDToTypeID('#Pxl'), 0.000000 );
    desc22.putUnitDouble( charIDToTypeID('Vrtc'), charIDToTypeID('#Pxl'), 0.000000 );
    desc21.putObject( charIDToTypeID('Ofst'), charIDToTypeID('Ofst'), desc22 );
    executeAction( charIDToTypeID('Plc '), desc21, DialogModes.NO );
}

function alignToCanvas(alignToCanvas, alignValue) {

	app.activeDocument.selection.selectAll();

	var s2t = function (s) {
		return app.stringIDToTypeID(s);
	};
	var descriptor = new ActionDescriptor();
	var reference = new ActionReference();
	reference.putEnumerated( s2t( "layer" ), s2t( "ordinal" ), s2t( "targetEnum" ));
	descriptor.putReference( s2t( "null" ), reference );
	descriptor.putEnumerated( s2t( "using" ), s2t( "alignDistributeSelector" ), s2t( alignValue ));
	descriptor.putBoolean( s2t( "alignToCanvas" ), alignToCanvas );
	executeAction( s2t( "align" ), descriptor, DialogModes.NO );

	app.activeDocument.selection.deselect();
}

function readPlist() {
	var plist = new File(PREFS_FOLDER + '/' + PREFS);
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
