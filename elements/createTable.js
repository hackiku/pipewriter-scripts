// elements/createTable.js - Programmatic table creation

// Expose global functions
var tableCreator = tableCreator || {};

/**
 * Create zigzag right element at cursor position
 * @returns {Object} Result object with success/error info
 */
tableCreator.createZigzagRight = function () {
	try {
		const doc = DocumentApp.getActiveDocument();
		const cursor = doc.getCursor();

		if (!cursor) {
			throw new Error('No cursor position found. Please click where you want to insert the element.');
		}

		// Create a new table with 1 row and 3 columns
		const table = doc.getBody().appendTable([['', '', '']]);

		// Set column widths - middle column is narrower
		table.setColumnWidth(0, 275); // First column
		table.setColumnWidth(1, 100); // Middle column (narrower)
		table.setColumnWidth(2, 275); // Last column

		// Set overall table properties
		table.setBorderWidth(0); // No visible borders
		table.setPaddingBottom(0);
		table.setPaddingTop(0);
		table.setPaddingLeft(0);
		table.setPaddingRight(0);

		// Get the cells for content positioning
		const leftCell = table.getCell(0, 0);
		const middleCell = table.getCell(0, 1);
		const rightCell = table.getCell(0, 2);

		// Clear any default content
		leftCell.clear();
		middleCell.clear();
		rightCell.clear();

		// Right content (zigzag right puts content in the right cell)
		// Add heading text 
		const rightHeading = rightCell.appendParagraph('Zigzag right');
		rightHeading.setHeading(DocumentApp.ParagraphHeading.HEADING2);
		rightHeading.setAlignment(DocumentApp.HorizontalAlignment.CENTER);

		// Add description text
		const rightParagraph = rightCell.appendParagraph('Paragraph 10-40 words.');
		rightParagraph.setHeading(DocumentApp.ParagraphHeading.HEADING6);
		rightParagraph.setAlignment(DocumentApp.HorizontalAlignment.CENTER);

		// Set cell alignments
		rightCell.setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER);
		middleCell.setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER);

		// Move the table to the cursor position
		const element = cursor.getElement();
		const parent = element.getParent();
		const cursorPosition = parent.getChildIndex(element);
		const body = doc.getBody();

		// Insert at cursor position based on context
		if (parent.getType() == DocumentApp.ElementType.TABLE_CELL) {
			parent.insertTable(cursorPosition + 1, table);
		} else {
			body.insertTable(cursorPosition + 1, table);
		}

		// Remove the original table from the end
		body.removeChild(table);

		// Position cursor after the inserted table
		try {
			const tableIndex = body.getChildIndex(body.findElement(DocumentApp.ElementType.TABLE, cursor).getElement());
			if (tableIndex < body.getNumChildren() - 1) {
				const nextElement = body.getChild(tableIndex + 1);
				doc.setCursor(doc.newPosition(nextElement, 0));
			} else {
				const newPara = body.appendParagraph('');
				doc.setCursor(doc.newPosition(newPara, 0));
			}
		} catch (e) {
			Logger.log('Error positioning cursor: ' + e);
		}

		return {
			success: true,
			message: 'Zigzag right element inserted at cursor position'
		};
	} catch (error) {
		Logger.log('Error inserting zigzag table: ' + error);
		return {
			success: false,
			error: error.message || 'Failed to insert zigzag table'
		};
	}
};