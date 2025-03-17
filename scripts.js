// scripts.gs


function changeBg(color) {
  var doc = DocumentApp.getActiveDocument();
  var body = doc.getBody();

  body.setBackgroundColor(color);
}

function invertTextColor(color) {
  // TODO make black text white, light gray text dark gray etc.
  // should apply to all HEADING types

}

function deleteColumnWithCursor() {
  var doc = DocumentApp.getActiveDocument();
  var cursor = doc.getCursor();

  if (cursor) {
    var element = cursor.getElement();
    var table = element.getAncestors().find(function(ancestor) {
      return ancestor.getType() === DocumentApp.ElementType.TABLE;
    });

    if (table) {
      var cell = element.asTableCell();
      var columnIndex = cell.getChildIndex();
      var numRows = table.asTable().getNumRows();

      for (var i = 0; i < numRows; i++) {
        table.asTable().getRow(i).removeCell(columnIndex);
      }
    } else {
      Logger.log('The cursor is not in a table.');
    }
  } else {
    Logger.log('Cannot find the cursor.');
  }
}



// ========================= html2doc =======================================


function html2doc() {

  const body = DocumentApp.getActiveDocument().getBody();
  const textAll = body.getText();
  const lines = textAll.split("\n");

  let storedText = "";

  // Loop through each line in the text
  for (let i = 0; i < lines.length; i++) {
    // current line
    const line = lines[i].trim();

    // Check if the line starts with an HTML tag
    if (line.startsWith("<")) {

      const tagName = line.substring(1, line.indexOf(">")).toLowerCase();
      const tagText = line.substring(line.indexOf(">") + 1, line.lastIndexOf("<")).trim();

      let paragraph = null;
      switch (tagName) {
        case "h2":
          // Add an empty paragraph before every Heading 2 paragraph
          if (storedText) {
            body.appendParagraph("").setHeading(DocumentApp.ParagraphHeading.NORMAL);
          }
          paragraph = body.appendParagraph(tagText).setHeading(DocumentApp.ParagraphHeading.HEADING2);
          break;
        case "h3":
          // Add an empty paragraph before every Heading 3 paragraph
          if (storedText) {
            body.appendParagraph("").setHeading(DocumentApp.ParagraphHeading.NORMAL);
          }
          paragraph = body.appendParagraph(tagText).setHeading(DocumentApp.ParagraphHeading.HEADING3);
          break;
        case "h4":
          paragraph = body.appendParagraph(tagText).setHeading(DocumentApp.ParagraphHeading.HEADING4);
          break;
        case "h5":
          paragraph = body.appendParagraph(tagText).setHeading(DocumentApp.ParagraphHeading.HEADING5);
          break;
        case "p":
          paragraph = body.appendParagraph(tagText).setHeading(DocumentApp.ParagraphHeading.HEADING6);
          break;
        default:
          // If the tag name is not recognized, ignore the line
          break;
      }

      // Add a new line after the paragraph
      if (paragraph) {
        storedText += line + "\n";
      }
    }
  }
}


// ========================= deleteHTMLtags =======================================


function deleteHTMLtags() {
  var doc = DocumentApp.getActiveDocument();
  var body = doc.getBody();
  
  var tags = ["<h2>", "</h2>", "<h3>", "</h3>", "<h4>", "</h4>", "<h5>", "</h5>", "<h6>", "</h6>", "<p>", "</p>"];

  for (var i = 0; i < tags.length; i += 2) {
    body.replaceText(tags[i] + "|" + tags[i + 1], "");
  }
}


// ========================= doc2html =======================================

function doc2html() {
  const body = DocumentApp.getActiveDocument().getBody();
  const paras = body.getParagraphs();

  // Define the prefix map
  const prefixMap = {
    [DocumentApp.ParagraphHeading.HEADING1]: '',
    [DocumentApp.ParagraphHeading.HEADING2]: '<h2>',
    [DocumentApp.ParagraphHeading.HEADING3]: '<h3>',
    [DocumentApp.ParagraphHeading.HEADING4]: '<h4>',
    [DocumentApp.ParagraphHeading.HEADING5]: '<h5>',
    [DocumentApp.ParagraphHeading.HEADING6]: '<p>'
  };

  // Define the suffix map
  const suffixMap = {
    [DocumentApp.ParagraphHeading.HEADING1]: '',
    [DocumentApp.ParagraphHeading.HEADING2]: '</h2>',
    [DocumentApp.ParagraphHeading.HEADING3]: '</h3>',
    [DocumentApp.ParagraphHeading.HEADING4]: '</h4>',
    [DocumentApp.ParagraphHeading.HEADING5]: '</h5>',
    [DocumentApp.ParagraphHeading.HEADING6]: '</p>'
  };

  // Initialize an empty array to store modified text
  let modifiedTextArr = [];

  // Loop through each paragraph in the document
  for (let i = 0; i < paras.length; i++) {
    // Get the current paragraph
    const para = paras[i];

    // Check if the paragraph is a heading paragraph (i.e., not of normal type)
    if (
      para.getType() === DocumentApp.ElementType.PARAGRAPH &&
      para.getHeading() !== DocumentApp.ParagraphHeading.NORMAL
    ) {
      // Get the level of the heading
      const level = para.getHeading();

      // Get the text of the heading
      const text = para.getText();

      // Check if the heading text is not empty
      if (text.trim() !== "") {
        // Format the text with prefix and suffix based on the heading level
        const formattedText = prefixMap[level] + text + suffixMap[level];

        // Store the formatted text along with the heading level
        modifiedTextArr.push({formattedText, level});
      }
    }
  }

  // Insert each line as a separate paragraph with NORMAL heading
  for (let i = 0; i < modifiedTextArr.length; i++) {
    let paragraph = body.insertParagraph(i, modifiedTextArr[i].formattedText).setHeading(DocumentApp.ParagraphHeading.NORMAL);

    // If the original paragraph was HEADING2 or HEADING3, bold the new paragraph
    if (modifiedTextArr[i].level === DocumentApp.ParagraphHeading.HEADING2 || 
        modifiedTextArr[i].level === DocumentApp.ParagraphHeading.HEADING3) {
      // Iterate over text elements in the paragraph and bold them
      let textElements = paragraph.editAsText();
      textElements.setBold(true);
    }
  }
}
