// indexDropper.js - Direct index-based element insertion
// Uses pre-mapped element indexes instead of scanning the document

// Expose global object for the add-on
var indexDropper = {};

// Master document IDs
const INDEX_DOCS = {
	light: "1yDx1RzvNqTHpPkdfml1iHn7H32a3kvvZqIm8NTmIKPk", // Dev/test light doc
	dark: "1tD1NcdqhEyTy3N6K2Syu4f6Pn-aDYET9GTlB1hlBjxM"  // Dev/test dark doc
};

// Pre-mapped element indexes for direct access
const ELEMENT_INDEX = {
	"container-center": 2,
	"background-empty": 6,
	"background-color": 10,
	"hero": 14,
	"zz-left": 18,
	"zz-right": 22,
	"placeholder": 26,
	"blurbs-3": 30,
	"blurbs-4": 34,
	"blurbs-vertical-3": 37,
	"list-1": 41,
	"list-2": 45,
	"list-3": 49,
	"button-primary-left": 54,
	"button-secondary-left": 58,
	"buttons-2-left": 62,
	"button-primary-center": 67,
	"button-secondary-center": 71,
	"buttons-2-center": 75,
	"cards-2": 79,
	"cards-3": 83,
	"cards-4": 87,
	"cards-2x2": 91,
	"cards-6": 95,
	"pricing-2": 99,
	"styleguide": 103
};

// Default theme constant
const INDEX_THEME = "light";

/**
 * Gets an element directly by index from the master document
 * @param {string} elementId - The ID of the element to retrieve
 * @param {string} theme - Theme ('light' or 'dark')
 * @returns {Table|null} The copied table element or null if not found
 */
function getElementByIndex(elementId, theme = INDEX_THEME) {
	try {
		// Check if element exists in the mapping
		if (!(elementId in ELEMENT_INDEX)) {
			throw new Error(`Element "${elementId}" not found in index mapping`);
		}

		// Get the element index
		const elementIndex = ELEMENT_INDEX[elementId];

		// Open document and get element directly by index
		const masterDoc = DocumentApp.openById(INDEX_DOCS[theme]);
		const masterBody = masterDoc.getBody();
		const element = masterBody.getChild(elementIndex);

		if (element && element.getType() == DocumentApp.ElementType.TABLE) {
			return element.copy();
		} else {
			throw new Error(`Element at index ${elementIndex} is not a table`);
		}
	} catch (error) {
		Logger.log('Failed to get element by index: ' + error);
		return null;
	}
}

/**
 * Enhanced table insertion with cursor positioning
 * @param {Table} table - The table to insert
 * @returns {Object|false} Result object or false on failure
 */
function insertIndexTable(table) {
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
				tableIndex: tableIndex,
				executionTime: null // Will be set by the caller
			};
		} catch (error) {
			Logger.log('Error positioning cursor: ' + error);
			return {
				success: true,
				cursorMoved: false,
				error: error.toString(),
				executionTime: null // Will be set by the caller
			};
		}
	}

	return false;
}

/**
 * Main element insertion function using direct index lookup
 * @param {string|Object} params - Element ID or parameter object
 * @returns {Object} Result object with success/error info
 */
function insertIndexElement(params) {
	const startTime = new Date().getTime();

	try {
		// Handle both the object form and direct string form
		const elementId = typeof params === 'object' ? params.elementId : params;
		const theme = (typeof params === 'object' && params.theme) ? params.theme : INDEX_THEME;

		Logger.log(`Getting indexed element: ${elementId} (${theme})`);

		// Get the table directly by index
		const table = getElementByIndex(elementId, theme);

		if (!table) {
			throw new Error(`Element ${elementId} not found or could not be copied`);
		}

		// Insert table and position cursor
		const result = insertIndexTable(table);
		if (!result) {
			throw new Error('Failed to insert table');
		}

		// Add execution time
		result.executionTime = new Date().getTime() - startTime;

		return {
			success: true,
			message: `Element ${elementId} inserted successfully`,
			cursorMoved: result.cursorMoved,
			tableIndex: result.tableIndex,
			executionTime: result.executionTime
		};

	} catch (error) {
		Logger.log('Error in insertIndexElement: ' + error);
		return {
			success: false,
			error: error.message || 'Failed to insert element',
			executionTime: new Date().getTime() - startTime
		};
	}
}

/**
 * Insert zigzag left element (menu helper)
 */
function insertZigzagLeft() {
	const result = insertIndexElement('zz-left');

	// Show execution time in alert for testing
	if (result.success) {
		DocumentApp.getUi().alert(`Zigzag Left inserted successfully!\nExecution time: ${result.executionTime}ms`);
	} else {
		DocumentApp.getUi().alert(`Error: ${result.error}\nExecution time: ${result.executionTime}ms`);
	}

	return result;
}

/**
 * Insert blurbs-3 element (menu helper)
 */
function insertBlurbs3() {
	const result = insertIndexElement('blurbs-3');

	// Show execution time in alert for testing
	if (result.success) {
		DocumentApp.getUi().alert(`Blurbs-3 inserted successfully!\nExecution time: ${result.executionTime}ms`);
	} else {
		DocumentApp.getUi().alert(`Error: ${result.error}\nExecution time: ${result.executionTime}ms`);
	}

	return result;
}

// Make functions available through the indexDropper global object
indexDropper.insertElement = insertIndexElement;
indexDropper.insertZigzagLeft = insertZigzagLeft;
indexDropper.insertBlurbs3 = insertBlurbs3;