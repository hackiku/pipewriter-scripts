// PIPE_BOUND.js
/**
 * working version used by users as bound script
 */

const USE_CACHE = false; // Toggle caching
const CACHE_TTL = 21600; // 6 hours cache duration

const MASTER_DOCS = {
	light: "1gVSTS5SLDuui85maXIVtOjXzg7aBbviZ_6XHXSLYE74",
	dark: "1FU1sZ4KdeAv_VcvDexzq6D4F0tffXnuVYAVVeVxz-ik"
};

// Helper to get element table from master doc
function getElementFromMaster(elementId, theme) {
	try {
		const masterDoc = DocumentApp.openById(MASTER_DOCS[theme]);
		const masterBody = masterDoc.getBody();
		let foundElement = false;
		let table = null;

		// Adjust elementId for dark theme
		// const adjustedElementId = theme === "dark" ? `${elementId}-dark` : elementId;
		const adjustedElementId = theme === "dark" ? `${elementId}` : elementId;

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
		const { elementId, theme = THEME } = params;
		Logger.log(`Getting element: ${elementId} (${theme})`);

		// Adjust elementId for dark theme
		const adjustedElementId = theme === "dark" ? `${elementId}-dark` : elementId;

		// Try cache first
		let table = getCachedElement(adjustedElementId, theme);

		// If not in cache, get from master doc
		if (!table) {
			table = getElementFromMaster(elementId, theme);
			if (!table) {
				throw new Error(`Element ${adjustedElementId} not found`);
			}
			// Cache the table
			cacheElement(adjustedElementId, theme, table);
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
			tableIndex: result.tableIndex,
			fromCache: !!getCachedElement(adjustedElementId, theme)
		};

	} catch (error) {
		Logger.log('Error in getElement: ' + error);
		return {
			success: false,
			error: error.message || 'Failed to insert element'
		};
	}
}

// Menu setup
function onOpen() {
	var ui = DocumentApp.getUi();
	ui.createMenu("Pipewriter")
		.addItem("Open App", "showFormInSidebar")
		.addSeparator()
		.addItem("HTML to start ↑", "dropHtmlStart")
		.addItem("HTML to end ↓", "dropHtmlEnd")
		.addItem("Strip HTML tags", "stripHtmlTags")
		.addItem("Delete HTML", "stripHtmlAll")
		.addItem("HTML to clipboard", "dropHtmlClipboard")
		.addToUi();
}

// Wrapper functions for menu items
function dropHtmlStart() {
	return dropHtml({ position: 'start' });
}
function dropHtmlEnd() {
	return dropHtml({ position: 'end' });
}
function dropHtmlClipboard() {
	return dropHtml({ copyToClipboard: true });
}

function showFormInSidebar() {
	var form = HtmlService.createTemplateFromFile("Index")
		.evaluate()
		.setTitle("Pipewriter");
	DocumentApp.getUi().showSidebar(form);
}