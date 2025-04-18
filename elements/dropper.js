// dropper.js - Streamlined index-based element insertion

// Expose global dropper object
var dropper = {};

// Master document IDs
const MASTER_DOCS = {
	light: "1gVSTS5SLDuui85maXIVtOjXzg7aBbviZ_6XHXSLYE74",
	dark: "1FU1sZ4KdeAv_VcvDexzq6D4F0tffXnuVYAVVeVxz-ik"
};

// Element indexes (light theme)
const LIGHT_ELEMENT_INDEX = {
	"container-center": 2, "background-empty": 6, "background-color": 10,
	"hero": 14, "zz-left": 18, "zz-right": 22, "placeholder": 26,
	"blurbs-3": 30, "blurbs-4": 34, "blurbs-vertical-3": 37,
	"list-1": 41, "list-2": 45, "list-3": 49,
	"button-primary-left": 54, "button-secondary-left": 58, "buttons-2-left": 62,
	"button-primary-center": 67, "button-secondary-center": 71, "buttons-2-center": 75,
	"cards-2": 79, "cards-3": 83, "cards-4": 87, "cards-2x2": 91,
	"cards-6": 95, "pricing-2": 99, "styleguide": 103
};

// Element indexes (dark theme)
const DARK_ELEMENT_INDEX = {
	"container-center": 2, "background-empty": 6, "background-color": 10,
	"hero": 14, "zz-left": 18, "zz-right": 26, "placeholder": 34,
	"blurbs-3": 38, "blurbs-4": 46, "blurbs-vertical-3": 78,
	"list-1": 54, "list-2": 62, "list-3": 70,
	"button-primary-left": 86, "button-secondary-left": 90, "buttons-2-left": 94,
	"button-primary-center": 98, "button-secondary-center": 102, "buttons-2-center": 106,
	"cards-2": 110, "cards-3": 118, "cards-4": 126, "cards-2x2": 135,
	"cards-6": 143, "pricing-2": 151, "styleguide": 159
};

// Document cache (keeps open document references)
const documentCache = {
	light: null,
	dark: null,
	// Clear cache if needed (e.g., after a long period)
	clear: function () {
		this.light = null;
		this.dark = null;
	}
};

/**
 * Get element table from the master document using direct index
 */
function getElementTable(elementId, theme = 'light') {
	try {
		// Get the right index map
		const indexMap = theme === 'light' ? LIGHT_ELEMENT_INDEX : DARK_ELEMENT_INDEX;

		// Verify element exists in index map
		if (!(elementId in indexMap)) {
			throw new Error(`Element "${elementId}" not found in ${theme} theme`);
		}

		// Get or open master document (using cache)
		let masterDoc = documentCache[theme];
		if (!masterDoc) {
			masterDoc = DocumentApp.openById(MASTER_DOCS[theme]);
			documentCache[theme] = masterDoc; // Cache for future use
		}

		// Get element directly by index
		const index = indexMap[elementId];
		const masterBody = masterDoc.getBody();

		if (index >= masterBody.getNumChildren()) {
			throw new Error(`Element index out of bounds: ${index}`);
		}

		const element = masterBody.getChild(index);

		// Verify it's a table
		if (!element || element.getType() !== DocumentApp.ElementType.TABLE) {
			throw new Error(`Element at index ${index} is not a table`);
		}

		return element.copy();
	} catch (error) {
		Logger.log(`Error getting element: ${error}`);
		return null;
	}
}

/**
 * Get cursor information and handle selection
 * Returns the cursor element and parent, or null if no valid cursor
 */
function getCursorInfo() {
	try {
		const doc = DocumentApp.getActiveDocument();

		// Check for cursor
		const cursor = doc.getCursor();

		// If cursor exists (blinking insertion point)
		if (cursor) {
			const element = cursor.getElement();
			const parent = element.getParent();
			const offset = parent.getChildIndex(element);

			return {
				type: 'cursor',
				element: element,
				parent: parent,
				offset: offset,
				selection: null
			};
		}

		// Check for selection
		const selection = doc.getSelection();
		if (selection) {
			const rangeElements = selection.getRangeElements();
			if (rangeElements && rangeElements.length > 0) {
				// Get the first element in the selection
				const firstElement = rangeElements[0].getElement();

				// Find a suitable insertion point
				const parent = firstElement.getParent();
				let referenceElement = firstElement;

				// If it's partial text selection, get the parent paragraph
				if (firstElement.getType() === DocumentApp.ElementType.TEXT) {
					referenceElement = parent;
				}

				// Get position in parent's parent (usually the body)
				const container = parent.getParent();
				const offset = container.getChildIndex(referenceElement);

				return {
					type: 'selection',
					element: firstElement,
					parent: container,
					offset: offset,
					selection: selection
				};
			}
		}

		// No cursor or selection found
		throw new Error('No cursor or selection found. Please click in the document where you want to insert the element.');

	} catch (error) {
		Logger.log(`Error getting cursor info: ${error}`);
		return null;
	}
}

/**
 * Insert table at cursor position and handle cursor repositioning
 */
function insertTableAtCursor(table) {
	if (!table) {
		throw new Error('No table to insert');
	}

	try {
		const doc = DocumentApp.getActiveDocument();
		const body = doc.getBody();
		const cursorInfo = getCursorInfo();

		if (!cursorInfo) {
			throw new Error('Please position your cursor in the document before inserting');
		}

		let insertedTable;

		// Handle selection vs cursor
		if (cursorInfo.type === 'selection') {
			// If there's a selection, replace it or insert at its position
			try {
				cursorInfo.selection.removeFromParent();
			} catch (e) {
				// Some selections can't be removed directly
				Logger.log("Couldn't remove selection: " + e);
			}
		}

		// Insert table
		if (cursorInfo.parent.getType() === DocumentApp.ElementType.TABLE_CELL) {
			// If inside a table cell
			insertedTable = cursorInfo.parent.insertTable(cursorInfo.offset + 1, table);
		} else {
			// Regular document body or other container
			insertedTable = body.insertTable(cursorInfo.offset + 1, table);
		}

		// Position cursor after table
		if (insertedTable) {
			try {
				const tableIndex = body.getChildIndex(insertedTable);

				if (tableIndex < body.getNumChildren() - 1) {
					// If there's an element after the table
					const nextElement = body.getChild(tableIndex + 1);
					doc.setCursor(doc.newPosition(nextElement, 0));
				} else {
					// If table is the last element, add paragraph and set cursor
					const newPara = body.appendParagraph('');
					doc.setCursor(doc.newPosition(newPara, 0));
				}

				return {
					success: true,
					insertedTable: insertedTable,
					tableIndex: tableIndex
				};
			} catch (e) {
				Logger.log('Error positioning cursor: ' + e);
				// Still consider insertion successful
				return {
					success: true,
					insertedTable: insertedTable,
					cursorError: e.toString()
				};
			}
		}

		throw new Error('Failed to insert table');
	} catch (error) {
		Logger.log('Table insertion error: ' + error);
		return {
			success: false,
			error: error.toString()
		};
	}
}

/**
 * Main element insertion function
 */
function getElement(params) {
	try {
		// Handle both string and object forms
		const elementId = typeof params === 'object' ? params.elementId : params;
		const theme = (typeof params === 'object' && params.theme) ? params.theme : 'light';

		// Get the element table
		const table = getElementTable(elementId, theme);
		if (!table) {
			throw new Error(`Could not retrieve element: ${elementId}`);
		}

		// Insert at cursor position
		const result = insertTableAtCursor(table);
		if (!result.success) {
			throw new Error(result.error || 'Failed to insert element');
		}

		return {
			success: true,
			message: `Element ${elementId} inserted successfully`,
			tableIndex: result.tableIndex
		};
	} catch (error) {
		return {
			success: false,
			error: error.message || 'Failed to insert element'
		};
	}
}

// Clear document cache periodically (every 10 minutes)
function scheduleDocumentCacheClearing() {
	try {
		const scriptProperties = PropertiesService.getScriptProperties();
		const lastCacheClearStr = scriptProperties.getProperty('lastCacheClear');
		const now = new Date().getTime();

		if (!lastCacheClearStr || now - parseInt(lastCacheClearStr) > 10 * 60 * 1000) {
			documentCache.clear();
			scriptProperties.setProperty('lastCacheClear', now.toString());
			Logger.log('Document cache cleared');
		}
	} catch (e) {
		// Fail silently
		Logger.log('Cache cleanup error: ' + e);
	}
}

// Set exports
dropper.getElement = getElement;