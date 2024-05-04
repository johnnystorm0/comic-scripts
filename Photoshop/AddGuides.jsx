/**
 * Add Guides
 * The Add Guides script will add guides to your Photoshop document that
 * represent the safe and trim areas.
 *
 * @author        Johnny Storm
 * @version        1.0.0
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
@@@BUILDINFO@@@ Add Guides to Page.jsx 1.0.0
*/

#target photoshop
#script 'Add Guides to Page';

const TRIM = .125;
const SAFE = .375;

if (app.documents.length > 0) {

    var originalRulerUnits = app.preferences.rulerUnits;
    app.preferences.rulerUnits = Units.INCHES;

    var doc = app.activeDocument;
    var width = doc.width.value;
    var height = doc.height.value;

    doc.guides.removeAll();
    doc.guides.add(Direction.HORIZONTAL, TRIM);
    doc.guides.add(Direction.HORIZONTAL, height - TRIM);
    doc.guides.add(Direction.HORIZONTAL, SAFE);
    doc.guides.add(Direction.HORIZONTAL, height - SAFE);

    doc.guides.add(Direction.VERTICAL, TRIM);
    doc.guides.add(Direction.VERTICAL, width - TRIM);
    doc.guides.add(Direction.VERTICAL, SAFE);
    doc.guides.add(Direction.VERTICAL, width - SAFE);

    app.preferences.rulerUnits = originalRulerUnits;
}
