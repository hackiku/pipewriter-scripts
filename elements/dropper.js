// elements/dropper.js - Production-aligned implementation

// Expose global dropper object from Code.js
var dropper = {};

// Define master document IDs - using the correct production IDs
const MASTER_DOCS = {
	light: "1gVSTS5SLDuui85maXIVtOjXzg7aBbviZ_6XHXSLYE74",
	dark: "1FU1sZ4KdeAv_VcvDexzq6D4F0tffXnuVYAVVeVxz-ik"
};

// Default theme constant
const THEME = "light";

// Helper to get element table from master document
function getElementFromMaster(elementId, theme) {
	try {
		const masterDoc = DocumentApp.openById(MASTER_DOCS[theme]);
		const masterBody = masterDoc.getBody();
		let foundElement = false;
		let table = null;

		// Note the difference in adjusting the elementId compared to your previous code
		// In production code, you keep the elementId as is and don't append '-dark'
		const adjustedElementId = elementId;

		// Find elementId paragraph and its table
		const numElements = masterBody.getNumChildren();
		for (let i = 0; i < numElements; i++) {
			const element = masterBody.getChild(i);

			if (element.getType() == DocumentApp.ElementType.PARAGRAPH) {
				if (element.getText().trim() === adjustedElementId) {
					foundElement = true;
					Logger.log('Found element: ' + adjustedElementId);
				}
			} else if (foundElement && element.getType() == DocumentApp.ElementType.TABLE) {
				table = element.copy();
				Logger.log('Found table for: ' + adjustedElementId);
				break;
			}
		}

		return table;

	} catch (error) {
		Logger.log('Failed to get element from master: ' + error);
		return null;
	}
}

// Enhanced table insertion with cursor positioning
function insertElementTable(table) {
	if (!table) {
		throw new Error('No table provided');
	}

	const doc = DocumentApp.getActiveDocument();
	const cursor = doc.getCursor();
	if (!cursor) {
		throw new Error('No cursor position found');
	}

	const element = cursor.getElement();
	const parent = element.getParent();
	const offset = parent.getChildIndex(element);

	let insertedTable;

	// Insert table based on context
	if (parent.getType() == DocumentApp.ElementType.TABLE_CELL) {
		insertedTable = parent.insertTable(offset + 1, table);
	} else {
		insertedTable = doc.getBody().insertTable(offset + 1, table);
	}

	// Position cursor after the inserted table
	if (insertedTable) {
		try {
			// Get index of inserted table
			const body = doc.getBody();
			const tableIndex = body.getChildIndex(insertedTable);

			// Try to insert and position cursor at a new paragraph after table
			let cursorPosition;

			if (tableIndex < body.getNumChildren() - 1) {
				// If there's an element after the table, position cursor there
				const nextElement = body.getChild(tableIndex + 1);
				cursorPosition = doc.newPosition(nextElement, 0);
			} else {
				// If table is last element, append paragraph and position cursor there
				const newPara = body.appendParagraph('');
				cursorPosition = doc.newPosition(newPara, 0);
			}

			// Set the new cursor position
			doc.setCursor(cursorPosition);

			Logger.log('Cursor positioned after table');
			return {
				success: true,
				cursorMoved: true,
				tableIndex: tableIndex
			};
		} catch (error) {
			Logger.log('Error positioning cursor: ' + error);
			return {
				success: true,
				cursorMoved: false,
				error: error.toString()
			};
		}
	}

	return false;
}

// Main element insertion function called by client
function getElement(params) {
	try {
		// Handle both the object form and direct string form
		const elementId = typeof params === 'object' ? params.elementId : params;
		const theme = (typeof params === 'object' && params.theme) ? params.theme : THEME;

		Logger.log(`Getting element: ${elementId} (${theme})`);

		// Adjust elementId for dark theme - crucial difference in your production code
		const adjustedElementId = theme === "dark" ? `${elementId}-dark` : elementId;

		// Get the table from master doc
		const table = getElementFromMaster(adjustedElementId, theme);

		if (!table) {
			throw new Error(`Element ${adjustedElementId} not found`);
		}

		// Insert table and position cursor
		const result = insertElementTable(table);
		if (!result) {
			throw new Error('Failed to insert table');
		}

		return {
			success: true,
			message: 'Element inserted successfully',
			cursorMoved: result.cursorMoved,
			tableIndex: result.tableIndex
		};

	} catch (error) {
		Logger.log('Error in getElement: ' + error);
		return {
			success: false,
			error: error.message || 'Failed to insert element'
		};
	}
}

// Make getElement available through the dropper global object
dropper.getElement = getElement;