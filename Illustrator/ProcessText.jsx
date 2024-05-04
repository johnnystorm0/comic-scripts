/**
 * Process Text
 * The Process Text script takes the pasted input from a comic book script and
 * processes the text to remove double spaces, converts quotes to smart quotes
 * and converts en and em dashes to double dashes (--). After the script finishes
 * processing the text it creates text objects for each paragraph.
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

/*
@@@BUILDINFO@@@ ProcessText.jsx 1.0.2.0
*/
#target illustrator
#script 'Process Comic Script';

app.preferences.setBooleanPreference("ShowExternalJSXWarning", false);

const VERSION    = "1.0.2";
const HYPEN     = 34;
const EMDASH    = 8212;
const ENDASH    = 8211;
const LDQUO     = 8220;
const RDQUO     = 8221;

Array.prototype.indexOf = function ( item ) {
    for (var i = 0 ; i < this.length; i++ ) {
        if ( this[i] == item ) {
            return i;
        }
    }
    return -1;
};

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

    var plist = new File(Folder.appData + '/' + PREFS);
        plist.open('w');
        plist.writeln(root.toXMLString());
        plist.close();

};

const OS = $.os.toLowerCase().indexOf('mac') >= 0 ? "MAC": "WINDOWS";
const PREFS     = 'com.johnnystorm.default.plist';
var prefs       = readPlist();

if (app.documents.length == 0) {
    alert('Please open a file prior to use.');
} else {
    var doc = app.activeDocument;

    const width      = doc.artboards[0].artboardRect[2];
    const height     = doc.artboards[0].artboardRect[3];
    const fonts      = findComicFonts();
    const fontNames  = findComicFontNames();
    var selectedFontIndex = prefs && prefs.selectedFont ? findFont(prefs.selectedFont) : 0;
    var selectedFont      = fonts[selectedFontIndex];
    const dlg = createDlg();
}

function createDlg() {
    var dlg                        = new Window('dialog', 'Prep Lettering Files');

    var body                       = dlg.add("panel {orientation: 'row'}");

    var row                        = body.add('group');
        row.orientation            = 'column';
        row.minimumSize.width      = 450;
        row.alignment              = [ ScriptUI.Alignment.LEFT, ScriptUI.Alignment.TOP];
        row.alignChildren          = ['left','bottom'];

    var label                      = row.add("Group {orientation: 'row', alignment: 'left'}");
        label.s                    = row.add("StaticText { text:'Select Dialog Font:', minimumSize: [75,0] }");

    var dd                         = row.add("DropDownList", undefined, fontNames);
        dd.items[ selectedFontIndex ].selected = true
    var text                       = row.add("Group {orientation: 'row', alignment: 'left'}");
        text.s                     = row.add("StaticText { text:'Paste Text Here:', minimumSize: [75,0] }");
    this.value                     = row.add('EditText', [0, 0, 450, 200], '', {multiline:true});

    var buttons                        = dlg.add('group');
        buttons.orientation            = 'row';

    buttons.left                  = buttons.add('group');
    buttons.left.orientation      = 'row';
    buttons.left.alignChildren    = 'left';
    buttons.left.minimumSize      = [300, 0];
    buttons.right                 = buttons.add('group');
    buttons.right.orientation     = 'row';
    buttons.right.alignChildren   = 'right';

    var version                   = buttons.left.add('statictext', undefined, 'Version: ' + VERSION, {readonly:true,multiline:false});
    var cancelBtn                 = buttons.right.add('button', undefined, 'Cancel', {name:'cancel'});
    var okBtn                     = buttons.right.add('button', undefined, 'OK', {name:'ok'});
        okBtn.onClick = function() {
            okBtn.enabled = false;
            dlg.close(1);
        }

    dlg.center();
    if (dlg.show() == 1) {
        var index = dd.selection.index;
        prefs.selectedFont = fontNames[index];
        prefs.writePlist();
        selectedFont = fonts[index];
        processText(this.value.text);
    }
}

function processText( str ) {
    var y = 0;
    var paragraphs = str.split('\n');
    for( var i=0; i < paragraphs.length; i++) {
        var paragraph = processParagraph(paragraphs[i]);
        if ( paragraph.length > 1 ) {
            var line = paragraph.pop();
            var speaker = paragraph[0];
            y = addTextFrame(speaker, y);
            y = addDialogue(line, y);
        } else {
            y = addTextFrame(paragraph[0], y);
        }

    }

}

function findFont(fontName) {
    return fontNames.indexOf( fontName ) > -1 ? fontNames.indexOf( fontName )  : 0;
}

function findComicFonts() {
    var fonts = [];

    for ( var i = 0; i < app.textFonts.length; i++ ) {
        var font = app.textFonts[i];

        if ( font.name.indexOf('BB') > -1 || font.name.indexOf('CC') > -1 ) {
            fonts.push(font);
        }
    }
    return fonts;
}

function findComicFontNames() {
    var fontNames = [];

    for ( var i = 0; i < app.textFonts.length; i++ ) {
        var font = app.textFonts[i];

        if ( font.name.indexOf('BB') > -1 || font.name.indexOf('CC') > -1 ) {
            fontNames.push(font.name);
        }
    }
    return fontNames;
}

function processParagraph( paragraph ) {
    var result = [];
    var re = new RegExp('(^[0-9.]+)([^:]*): ');
    paragraph = paragraph.replace( re, '$1 $2:\t');
    if (paragraph.indexOf('\t') > -1) {
        return paragraph.split('\t');
    } else {
        result.push(paragraph);
    }
    return result;
}

function addTextFrame( str, y ) {
    var textFrame = doc.textFrames.add();
        textFrame.contents = str;
        textFrame.top = y;
        textFrame.left = width + 25;

    return y - textFrame.height;
}

function addDialogue( str, y ) {
    var textFrame = doc.textFrames.add();
    var text = '';
    var start = true;
    str = str.replace('  ', ' ');

    for (var i = 0; i < str.length; i++) {
        switch ( str.charCodeAt(i) ) {
            case HYPEN:
                text += start ? String.fromCharCode(LDQUO) : String.fromCharCode(RDQUO);
                start = !start;
                break;
            case ENDASH:
                text += String.fromCharCode(45, 45);
                break;
            case EMDASH:
                text += String.fromCharCode(45, 45);
                break;
            default:
                text += str.charAt(i);
        }
    }
    textFrame.contents = text;
    var textArtRange = textFrame.textRange;
    textArtRange.characterAttributes.textFont = selectedFont;

    textFrame.top = y;
    textFrame.left = width + 25;

    return y - textFrame.height;
}


function readPlist() {
    var plist = new File(Folder.appData + '/' + PREFS);
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
