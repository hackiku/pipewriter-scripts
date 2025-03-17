// 4. convert wireframe to ai-friendly text

function aiReadDoc() {
  var doc = DocumentApp.getActiveDocument();
  var cursor = doc.getCursor();
  var body = doc.getBody();
  var cursorPosition = body.getChildIndex(cursor.getElement());

  // special naming convention for text to prepend
  var prependHeading = ['','h2: ', 'h3: ', 'feature or button: ', 'eyebrow: ', 'p: ']; 
  var prependHtml = ['','h2','h3','h4','h5','p'];

  // returned as object
  var readText = ""; //plain text output
  var readUx = ""; //store the format `h2: Headline` etc.
  var readHtml = ""; //store the format `<h2>Headline</h2> etc.

  // Manually map paragraph headings to integer equivalents
  var headingMap = {
    'HEADING1': 0,
    'HEADING2': 1,
    'HEADING3': 2,
    'HEADING4': 3,
    'HEADING5': 4,
    'HEADING6': 5
  }

  // Loop through all elements up to the cursor position
  for (var i = 0; i <= cursorPosition; i++) {
    var element = body.getChild(i);
    
    // If element is a paragraph
    if (element.getType() === DocumentApp.ElementType.PARAGRAPH) {
      var paragraph = element.asParagraph();
      if (paragraph.getHeading() !== DocumentApp.ParagraphHeading.NORMAL) {
        var heading = paragraph.getHeading(); // Get the heading level
        var headingIndex = headingMap[heading.toString()]; // Convert the heading level to an integer

        var text = paragraph.getText(); // Get the plain text

        readText += text + "\n"; // Append the plain text to readText
        readUx += prependHeading[headingIndex] + text + "\n"; // Append the formatted text to readUx
        readHtml += "<" + prependHtml[headingIndex] + ">" + text + "</" + prependHtml[headingIndex] + ">\n";
      }
    } 
    // If the element is a table, process it as before
    else if (element.getType() === DocumentApp.ElementType.TABLE) {
      var table = element.asTable();
      var numRows = table.getNumRows();
      // console log "table has X rows"
      for (var j = 0; j < numRows; j++) {
        var row = table.getRow(j);
        var numCells = row.getNumCells();
        for (var k = 0; k < numCells; k++) {
          var cell = row.getCell(k);
          var numChildren = cell.getNumChildren();
          for (var l = 0; l < numChildren; l++) {
            var child = cell.getChild(l);
            if (child.getType() === DocumentApp.ElementType.PARAGRAPH) {
              var paragraph = child.asParagraph();
              if (paragraph.getHeading() !== DocumentApp.ParagraphHeading.NORMAL) {
                
                var heading = paragraph.getHeading(); // Get the heading level
                var headingIndex = headingMap[heading.toString()]; // Convert the heading level to an integer

                var text = paragraph.getText(); // Get the plain text

                readText += text + "\n"; // Append the plain text to readText
                readUx += prependHeading[headingIndex] + text + "\n"; // Append the formatted text to readUx
                readHtml += "<" + prependHtml[headingIndex] + ">" + text + "</" + prependHtml[headingIndex] + ">\n"; // Append the HTML formatted text to readHtml
              }
            }
          }
        }
      }
    }
  }

  // return readText, readUx, readHtml
  return { readText: readText, readUx: readUx, readHtml: readHtml };
}