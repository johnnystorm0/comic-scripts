/**
 * Lettering Files
 *
 *
 * @author        Johnny Storm
 * @version        1.0.2
 *
 * 2024-04-16: Initial Version.
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

const VERSION        = "1.0.2";
const OS             = $.os.toLowerCase().indexOf('mac') >= 0 ? "MAC": "WINDOWS";
const PREFS          = 'com.johnnystorm.default.plist';
const DEFAULT_FOLDER = Folder.desktop;

var prefs            = readPlist();
var defaultPath      = prefs.defaultPath ? prefs.defaultPath : DEFAULT_FOLDER;
var outputPath       = prefs.outputPath ? prefs.outputPath : defaultPath;

var WIDTH            = 6.625 * 72;
var HEIGHT           = 10.1875 * 72;

const BLEED          = .125 * 72;
const safe           = .25;

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

createDlg();

function createDlg() {
    var dlg                        = new Window('dialog', 'Prep Lettering Files');

    var body                       = dlg.add("panel {orientation: 'row'}");

    var row                        = body.add('group');
        row.orientation            = 'column';
        row.minimumSize.width      = 450;
        row.alignment              = [ ScriptUI.Alignment.LEFT, ScriptUI.Alignment.TOP];
        row.alignChildren          = ['left','bottom'];

    var label                      = row.add("Group {orientation: 'row', alignment: 'left'}");
        label.s                    = row.add("StaticText { text:'Select Source Folder:', minimumSize: [75,0] }");

    var sourceButton               = row.add('button', undefined, 'Open Folder...', {name:'open1'});
    var sourceUri                  = row.add('statictext', [0,0,400,20], 'No files selected', {readonly:true,multiline:false});

    sourceButton.onClick    = function() {
        var folder = new Folder();
        if(sourceFolder = folder.selectDlg(undefined, defaultPath)) {
            prefs.defaultPath = sourceFolder.path + '/' + sourceFolder.name;
            prefs.writePlist();
            sourceUri.text = File.decode( prefs.defaultPath );
            defaultPath = prefs.defaultPath;
        }
    }

    var label2                     = row.add("Group {orientation: 'row', alignment: 'left'}");
        label2.s                   = row.add("StaticText { text:'Select Output Folder:', minimumSize: [75,0] }");
    var outputButton               = row.add('button', undefined, 'Open Folder...', {name:'open1'});
    var outputUri                  = row.add('statictext', [0,0,400,20], 'No files selected', {readonly:true,multiline:false});

    outputButton.onClick    = function() {
        var folder = new Folder();
        if(outputFolder = folder.selectDlg(undefined, outputPath)) {
            prefs.outputPath = outputFolder.path + '/' + outputFolder.name;
            prefs.writePlist();
            outputUri.text = File.decode( prefs.outputPath );
            outputPath = prefs.outputPath;
        }
    }

    var buttons                     = dlg.add('group');
        buttons.orientation         = 'row';
        buttons.left                = buttons.add('group');
        buttons.left.orientation    = 'row';
        buttons.left.alignChildren  = 'left';
        buttons.left.minimumSize    = [300, 0];
        buttons.right               = buttons.add('group');
        buttons.right.orientation   = 'row';
        buttons.right.alignChildren = 'right';

    var version                     = buttons.left.add('statictext', undefined, 'Version: ' + VERSION, {readonly:true,multiline:false});
    var cancelBtn                   = buttons.right.add('button', undefined, 'Cancel', {name:'cancel'});
    var okBtn                       = buttons.right.add('button', undefined, 'OK', {name:'ok'});
        okBtn.onClick = function() {
            if (sourceFolder != null) {
                okBtn.enabled = false;
                dlg.close(1);

            } else {
                alert('Please select a folder to process.');
            }
        }

    dlg.center();
    if (dlg.show() == 1) {
        prepPages(sourceFolder);
    }
}

function prepPages(sourceFolder) {
    createTempScript();
    var files = sourceFolder.getFiles("*.tif");
        files.sort();
    try {
        var bounds = getImageSize(files[0]);
        WIDTH  = bounds[2];
        HEIGHT = bounds[3];
        var pages = files.length;
        progress(pages);
        for(var i = 0; i < pages; i++) {
            progress.message("Processing page " + (i+1) + " of " + pages);
            var file = files[i];
            createPage(file.name, file);
            progress.increment();
        }
        progress.close();
        deleteTempScript();
        alert("Your files have been created.");
    } catch (e) {
        alert(e);
    }
}

function getImageSize(file) {
    var docPreset           = new DocumentPreset;
        docPreset.colorMode = DocumentColorSpace.CMYK;
        docPreset.title     = 'Untitled';
        docPreset.units     = RulerUnits.Inches;
        docPreset.width     = WIDTH;
        docPreset.height    = HEIGHT;

    var presetArt = app.startupPresetsList[3];
    var doc = app.documents.addDocument(presetArt, docPreset);
    var image = doc.placedItems.add();
        image.file = file;
    var w = image.width - (BLEED*2);
    var h = image.height - (BLEED*2);
    doc.close(SaveOptions.DONOTSAVECHANGES);
    return [0, 0, w, h];
}

function createDoc( filename, width, height ) {

    var docPreset           = new DocumentPreset;
        docPreset.colorMode = DocumentColorSpace.CMYK;
        docPreset.title     = filename;
        docPreset.units     = RulerUnits.Inches;
        docPreset.width     = width;
        docPreset.height    = height;

    // Startup Preset Options:
    //
    // 0 - Print
    // 1 - Film & Video
    // 2 - Web
    // 3 - Art & Illustration
    // 4 - Mobile
    // 5 - Film and Video
    var presetArt = app.startupPresetsList[3];

    return app.documents.addDocument(presetArt, docPreset);
}

function createPage(title, file) {
    var temp = title.split('.');
        temp.pop();
    var filename = temp.join('');

    var doc = createDoc( title, WIDTH, HEIGHT );

    var artwork = doc.layers[0]
        artwork.name = 'Artwork';
        artwork.locked = false;

    var image = doc.placedItems.add();
        image.file = file;
        image.position = [-BLEED, HEIGHT + BLEED];

    var w = image.width - (BLEED*2);
    var h = image.height - (BLEED*2);

    if ( WIDTH != w || HEIGHT != h ) {
        doc.close(SaveOptions.DONOTSAVECHANGES);
        doc = createDoc( title, w, h );
        artwork = doc.layers[0]
        artwork.name = 'Artwork';
        artwork.locked = false;

        image = doc.placedItems.add();
        image.file = file;
        image.position = [-BLEED,HEIGHT + BLEED];
    }

    setTemplateLayer();
    artwork.name   = 'Artwork';
    artwork.locked = true;

    var bleed = doc.layers.add();
        bleed.name = 'Bleed';

    var bleedBox = doc.pathItems.rectangle(-BLEED, -BLEED, image.width, -image.height);
        bleedBox.filled = false;
        bleedBox.strokeWidth = .5;

        bleed.locked = true;

    var bottom = doc.layers.add();
        bottom.name = 'Bottom Balloons';

    var tails = doc.layers.add();
        tails.name = 'Tails';

    var top = doc.layers.add();
        top.name = 'Top Balloons';

    var lettering = doc.layers.add();
        lettering.name = 'Lettering';

    var outputFile = File(outputPath + '/' + filename + '-lettering.ai');

    doc.saveAs(outputFile);
    doc.close();
}

function setTemplateLayer() {
    app.doScript ("setLayTmplAttr", "setLayTmplAttr", false); // action name, set name
}

function deleteTempScript() {
    app.unloadAction ("setLayTmplAttr", ""); // set name
}

function createTempScript () {
    var str = '/version 3'
        + '/name [ 14 7365744c6179546d706c41747472 ]'
        + '/isOpen 1'
        + '/actionCount 1'
        + '/action-1 { '
        + '/name [ 14 7365744c6179546d706c41747472 ] '
        + '/keyIndex 0 '
        + '/colorIndex 0 '
        + '/isOpen 1 '
        + '/eventCount 1 '
        + '/event-1 { '
        + '/useRulersIn1stQuadrant 0 '
        + '/internalName (ai_plugin_Layer) '
        + '/localizedName [ 5 4c61796572 ] '
        + '/isOpen 1 '
        + '/isOn 1 '
        + '/hasDialog 1 '
        + '/showDialog 0 '
        + '/parameterCount 10 '
        + '/parameter-1 { '
        + '/key 1836411236 '
        + '/showInPalette -1 '
        + '/type (integer) '
        + '/value 4 } '
        + '/parameter-2 { '
        + '/key 1851878757 '
        + '/showInPalette -1 '
        + '/type (ustring) '
        + '/value [ 20 4c61796572732050616e656c204f7074696f6e73 ] } '
        + '/parameter-3 { '
        + '/key 1953068140 '
        + '/showInPalette -1 '
        + '/type (ustring) '
        + '/value [ 7 4c617965722031 ] } '
        + '/parameter-4 { '
        + '/key 1953329260 '
        + '/showInPalette -1 '
        + '/type (boolean) '
        + '/value 1 } '
        + '/parameter-5 { '
        + '/key 1936224119 '
        + '/showInPalette -1 '
        + '/type (boolean) '
        + '/value 1 } '
        + '/parameter-6 { '
        + '/key 1819239275 '
        + '/showInPalette -1 '
        + '/type (boolean) '
        + '/value 1 } '
        + '/parameter-7 { '
        + '/key 1886549623 '
        + '/showInPalette -1 '
        + '/type (boolean) '
        + '/value 1 } '
        + '/parameter-8 { '
        + '/key 1886547572 '
        + '/showInPalette -1 '
        + '/type (boolean) '
        + '/value 0 } '
        + '/parameter-9 { '
        + '/key 1684630830 '
        + '/showInPalette -1 '
        + '/type (boolean) '
        + '/value 1 } '
        + '/parameter-10 { '
        + '/key 1885564532 '
        + '/showInPalette -1 '
        + '/type (unit real) '
        + '/value 50.0 '
        + '/unit 592474723 } }}'

    var f = new File ('~/ScriptAction.aia');

    f.open ('w');
    f.write (str);
    f.close ();
    app.loadAction (f);
    f.remove ();
}

function progress(steps) {
    var w = new Window("palette", "Progress", undefined, {closeButton: false});
    var t = w.add("statictext");
        t.preferredSize = [450, -1]; // 450 pixels wide, default height.

    if (steps) {
        var b = w.add("progressbar", undefined, 0, steps);
        b.preferredSize = [450, -1]; // 450 pixels wide, default height.
    }

    progress.close = function () {
        w.close();
    };

    progress.increment = function () {
        b.value++;
    };

    progress.message = function (message) {
        t.text = message;
    };

    w.show();
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
