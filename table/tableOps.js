// tableOps.gs - Enhanced version with better error handling and logging

/**
 * Handles table-related operations in Google Docs
 * @param {Object} params - Parameters passed from the frontend
 * @param {string} params.action - The action to perform
 * @param {Object} params.payload - The payload containing alignment values
 */
function tableOps(params) {
	const startTime = Date.now();

	try {
		console.log('tableOps called with params:', JSON.stringify(params));

		const doc = DocumentApp.getActiveDocument();
		const cursor = doc.getCursor();

		if (!cursor) {
			throw new Error('No cursor found in document. Please place your cursor in a table cell.');
		}

		// Get the current element and find the table cell
		let element = cursor.getElement();
		if (!element) {
			throw new Error('No element found at cursor position. Please place your cursor in a table cell.');
		}

		console.log('Starting element type:', element.getType());

		// Traverse up the element tree to find the table cell
		let attempts = 0;
		const maxAttempts = 10; // Prevent infinite loops

		while (element && element.getType() !== DocumentApp.ElementType.TABLE_CELL && attempts < maxAttempts) {
			element = element.getParent();
			attempts++;
			if (element) {
				console.log('Parent element type:', element.getType());
			}
		}

		if (!element || element.getType() !== DocumentApp.ElementType.TABLE_CELL) {
			throw new Error('Cursor is not inside a table cell. Please place your cursor in a table cell and try again.');
		}

		const tableCell = element;
		console.log('Found table cell, performing action:', params.action);

		switch (params.action) {
			case 'tableAlignHorizontal':
				return handleHorizontalAlignment(tableCell, params.payload, startTime);

			case 'tableAlignVertical':
				return handleVerticalAlignment(tableCell, params.payload, startTime);

			default:
				throw new Error(`Unknown table operation: ${params.action}`);
		}

	} catch (error) {
		console.error('Error in tableOps:', error);
		return {
			success: false,
			error: error.toString(),
			message: error.message || 'An unexpected error occurred',
			executionTime: Date.now() - startTime
		};
	}
}

/**
 * Handle horizontal table alignment
 * @param {GoogleAppsScript.Document.TableCell} tableCell 
 * @param {Object} payload 
 * @param {number} startTime 
 */
function handleHorizontalAlignment(tableCell, payload, startTime) {
	const hAlignmentMap = {
		'left': DocumentApp.HorizontalAlignment.LEFT,
		'center': DocumentApp.HorizontalAlignment.CENTER,
		'right': DocumentApp.HorizontalAlignment.RIGHT
	};

	if (!payload.alignment || !hAlignmentMap[payload.alignment]) {
		throw new Error(`Invalid horizontal alignment value: ${payload.alignment}. Must be 'left', 'center', or 'right'.`);
	}

	// Get the table from the cell
	const table = tableCell.getParent().getParent(); // cell -> row -> table
	if (table.getType() !== DocumentApp.ElementType.TABLE) {
		throw new Error('Could not find table element');
	}

	console.log(`Setting table horizontal alignment to: ${payload.alignment}`);

	// Apply horizontal alignment to the table
	table.setHorizontalAlignment(hAlignmentMap[payload.alignment]);

	return {
		success: true,
		message: `Table aligned ${payload.alignment}`,
		executionTime: Date.now() - startTime,
		details: `Successfully set table horizontal alignment to ${payload.alignment}`
	};
}

/**
 * Handle vertical cell content alignment
 * @param {GoogleAppsScript.Document.TableCell} tableCell 
 * @param {Object} payload 
 * @param {number} startTime 
 */
function handleVerticalAlignment(tableCell, payload, startTime) {
	const vAlignmentMap = {
		'top': DocumentApp.VerticalAlignment.TOP,
		'middle': DocumentApp.VerticalAlignment.CENTER,
		'bottom': DocumentApp.VerticalAlignment.BOTTOM
	};

	if (!payload.alignment || !vAlignmentMap[payload.alignment]) {
		throw new Error(`Invalid vertical alignment value: ${payload.alignment}. Must be 'top', 'middle', or 'bottom'.`);
	}

	console.log(`Setting cell vertical alignment to: ${payload.alignment}`);

	// Apply vertical alignment to the current cell
	tableCell.setVerticalAlignment(vAlignmentMap[payload.alignment]);

	return {
		success: true,
		message: `Cell content aligned ${payload.alignment}`,
		executionTime: Date.now() - startTime,
		details: `Successfully set cell vertical alignment to ${payload.alignment}`
	};
}

/**
 * Test function to verify table operations work
 */
function testTableOps() {
	console.log('Testing table operations...');

	// Test horizontal alignment
	const testParamsHorizontal = {
		action: 'tableAlignHorizontal',
		payload: {
			alignment: 'center'
		}
	};

	console.log('Testing horizontal alignment...');
	const resultHorizontal = tableOps(testParamsHorizontal);
	console.log('Horizontal result:', resultHorizontal);

	// Test vertical alignment
	const testParamsVertical = {
		action: 'tableAlignVertical',
		payload: {
			alignment: 'middle'
		}
	};

	console.log('Testing vertical alignment...');
	const resultVertical = tableOps(testParamsVertical);
	console.log('Vertical result:', resultVertical);

	return {
		horizontal: resultHorizontal,
		vertical: resultVertical
	};
}

/**
 * Helper function to get current table properties (for future use)
 */
function getCurrentTableProperties() {
	try {
		const doc = DocumentApp.getActiveDocument();
		const cursor = doc.getCursor();

		if (!cursor) {
			throw new Error('No cursor found');
		}

		let element = cursor.getElement();
		while (element && element.getType() !== DocumentApp.ElementType.TABLE_CELL) {
			element = element.getParent();
		}

		if (!element) {
			throw new Error('Not in a table cell');
		}

		const tableCell = element;
		const table = tableCell.getParent().getParent();

		return {
			success: true,
			data: {
				horizontalAlignment: table.getHorizontalAlignment(),
				verticalAlignment: tableCell.getVerticalAlignment(),
				// Add more properties as needed
			}
		};

	} catch (error) {
		return {
			success: false,
			error: error.toString()
		};
	}
}