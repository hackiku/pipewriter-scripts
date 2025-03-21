// elements/cursor.js - Simple table insertion at cursor position

// Expose global object with var to avoid initialization errors
var cursorTools = {};

/**
 * Get the current cursor position in the document
 * @returns {Object} Cursor information or null if no cursor found
 */
cursorTools.getCursorPosition = function () {
	try {
		const doc = DocumentApp.getActiveDocument();
		const cursor = doc.getCursor();

		if (!cursor) {
			return null;
		}

		return {
			element: cursor.getElement(),
			offset: cursor.getOffset(),
			surroundingText: cursor.getSurroundingText().getText()
		};
	} catch (error) {
		Logger.log('Error getting cursor position: ' + error);
		return null;
	}
};

/**
 * Insert a simple empty table at cursor position
 * @returns {Object} Result object with success/error info
 */
cursorTools.insertSimpleTable = function () {
	try {
		const doc = DocumentApp.getActiveDocument();
		const cursor = doc.getCursor();

		if (!cursor) {
			throw new Error('No cursor position found. Please click where you want to insert the table.');
		}

		// Create a simple 1x1 table
		const table = doc.getBody().appendTable([['']]);

		// Get cursor element and its parent
		const element = cursor.getElement();
		const parent = element.getParent();
		const cursorPosition = parent.getChildIndex(element);
		const body = doc.getBody();

		// Insert table at cursor position
		if (parent.getType() == DocumentApp.ElementType.TABLE_CELL) {
			// If cursor is in a table cell, insert table in that cell
			parent.insertTable(cursorPosition + 1, table);
		} else {
			// Otherwise insert in the body
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
			message: 'Simple table inserted at cursor position'
		};
	} catch (error) {
		Logger.log('Error inserting simple table: ' + error);
		return {
			success: false,
			error: error.message || 'Failed to insert simple table'
		};
	}
};