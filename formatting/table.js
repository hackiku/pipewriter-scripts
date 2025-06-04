// formatting/table.js

/**
 * Main table operations function
 * @param {Object} params - Operation parameters.
 * @param {string} params.action - The action to perform (e.g., 'setCellAlignment', 'setCellPadding', 'setBorders', 'selectWholeTable').
 * @param {string} [params.scope] - Scope for operations ('cell', 'table'). Required for alignment, padding, borders.
 * @param {string} [params.alignment] - Vertical alignment value ('top', 'middle', 'bottom') for setCellAlignment.
 * @param {number} [params.padding] - Padding value in points for setCellPadding.
 * @param {number} [params.borderWidth] - Border width in points for setBorders. Use 0 to remove borders.
 * @param {string} [params.borderColor] - Border color hex string (e.g., '#000000') for setBorders. Defaults to black if width > 0.
 * @param {boolean} [params.showAlert=true] - Whether to show UI alerts. Defaults to true. Set to false for appHandler calls.
 * @returns {Object} Result object with success/error info.
 */
function tableOps(params) {
	const startTime = new Date().getTime();
	// Default showAlert to true if not provided
	const showAlert = (params.showAlert === undefined) ? true : params.showAlert;
	let uiAlertMessage = null;

	try {
		if (!params || !params.action) {
			throw new Error('No action specified for tableOps');
		}

		// getTableContext might not be needed for all actions, or might have different requirements
		// For example, selectWholeTable might work if a table is selected, not just cursor in cell.
		const context = getTableContext(params.action === 'selectWholeTable');
		if (!context.success) {
			throw new Error(context.error);
		}

		let result;
		switch (params.action) {
			case 'setCellAlignment':
				result = setCellAlignment(context, params.scope, params.alignment, startTime);
				uiAlertMessage = result.message;
				break;

			case 'setCellPadding':
				result = setCellPadding(context, params.scope, params.padding, startTime);
				uiAlertMessage = result.message;
				break;

			case 'setBorders':
				result = setBorders(context, params.scope, params.borderWidth, params.borderColor, startTime);
				uiAlertMessage = result.message;
				break;

			case 'selectWholeTable':
				result = selectWholeTable(context, startTime);
				// No UI alert for selection, it's a visual change.
				break;

			default:
				throw new Error(`Unknown table action: ${params.action}`);
		}

		if (result.success && uiAlertMessage && showAlert) {
			DocumentApp.getUi().alert('✅ Table Control', uiAlertMessage, DocumentApp.getUi().ButtonSet.OK);
		}
		return result;

	} catch (error) {
		const errorResult = {
			success: false,
			error: error.toString(),
			message: error.message || 'An error occurred. Please ensure your cursor is inside a table cell for most operations.',
			executionTime: new Date().getTime() - startTime
		};
		Logger.log(`Error in tableOps (action: ${params.action}): ${error.message}\n${error.stack ? error.stack : ''}`);
		if (showAlert) {
			DocumentApp.getUi().alert('❌ Table Control Error', errorResult.message, DocumentApp.getUi().ButtonSet.OK);
		}
		return errorResult;
	}
}

/**
 * Get table context (table, cell, cursor info).
 * @param {boolean} allowTableSelectionOnly - If true, allows context even if only a table is selected (no cursor in cell).
 * @returns {Object} Context object or error.
 */
function getTableContext(allowTableSelectionOnly = false) {
	try {
		const doc = DocumentApp.getActiveDocument();
		const cursor = doc.getCursor();
		const selection = doc.getSelection();

		let table, tableCell = null, tableRow = null;

		if (cursor) {
			let element = cursor.getElement();
			if (!element) return { success: false, error: 'No element found at cursor position.' };

			let currentElement = element;
			let attempts = 0;
			const maxAttempts = 10;

			while (currentElement && currentElement.getType() !== DocumentApp.ElementType.TABLE_CELL && attempts < maxAttempts) {
				currentElement = currentElement.getParent();
				attempts++;
			}

			if (currentElement && currentElement.getType() === DocumentApp.ElementType.TABLE_CELL) {
				tableCell = currentElement.asTableCell();
				tableRow = tableCell.getParent().asTableRow();
				table = tableRow.getParent().asTable();
			}
		}

		// If no table found via cursor, or if allowed, check selection
		if ((!table || allowTableSelectionOnly) && selection) {
			const rangeElements = selection.getRangeElements();
			if (rangeElements.length === 1 && rangeElements[0].getElement().getType() === DocumentApp.ElementType.TABLE) {
				const selectedTable = rangeElements[0].getElement().asTable();
				if (!table) table = selectedTable; // Prioritize cursor-based table, but use selected if no cursor one
				// If table was already found by cursor, and selection is the same table, it's fine.
				// If different, it's ambiguous, but usually cursor is more specific.
			}
		}

		if (!table) {
			return { success: false, error: 'Cursor is not inside a table cell, nor is a table selected. Please click inside a table or select a table.' };
		}

		return {
			success: true,
			table: table,
			tableCell: tableCell, // Can be null if table was found via selection
			tableRow: tableRow,   // Can be null
			doc: doc
		};

	} catch (error) {
		Logger.log('Error in getTableContext: ' + error.message);
		return { success: false, error: 'Failed to get table context: ' + error.message };
	}
}

/**
 * Set cell content vertical alignment.
 * @param {Object} context - Table context.
 * @param {string} scope - 'cell' or 'table'.
 * @param {string} alignment - 'top', 'middle', or 'bottom'.
 * @param {number} startTime - Start time for execution tracking.
 * @returns {Object} Result.
 */
function setCellAlignment(context, scope, alignment, startTime) {
	if (!context || !context.table) throw new Error('Invalid table context for setCellAlignment.');
	if (scope === 'cell' && !context.tableCell) throw new Error('A cell must be active (cursor inside) to align only that cell.');
	if (!['top', 'middle', 'bottom'].includes(alignment)) throw new Error(`Invalid cell alignment: ${alignment}.`);
	if (!['cell', 'table'].includes(scope)) throw new Error(`Invalid scope for cell alignment: ${scope}.`);

	const alignmentMap = {
		'top': DocumentApp.VerticalAlignment.TOP,
		'middle': DocumentApp.VerticalAlignment.CENTER,
		'bottom': DocumentApp.VerticalAlignment.BOTTOM
	};
	const verticalAlignment = alignmentMap[alignment];
	let cellsUpdated = 0;

	if (scope === 'cell') {
		context.tableCell.setVerticalAlignment(verticalAlignment);
		cellsUpdated = 1;
	} else { // scope === 'table'
		const numRows = context.table.getNumRows();
		for (let i = 0; i < numRows; i++) {
			const row = context.table.getRow(i);
			const numCells = row.getNumCells();
			for (let j = 0; j < numCells; j++) {
				row.getCell(j).setVerticalAlignment(verticalAlignment);
				cellsUpdated++;
			}
		}
	}

	const scopeText = scope === 'cell' ? 'selected cell' : `all ${cellsUpdated} cells in the table`;
	return {
		success: true,
		message: `Content aligned to ${alignment} for ${scopeText}.`,
		executionTime: new Date().getTime() - startTime
	};
}

/**
 * Set cell padding.
 * @param {Object} context - Table context.
 * @param {string} scope - 'cell' or 'table'.
 * @param {number} padding - Padding in points.
 * @param {number} startTime - Start time for execution tracking.
 * @returns {Object} Result.
 */
function setCellPadding(context, scope, padding, startTime) {
	if (!context || !context.table) throw new Error('Invalid table context for setCellPadding.');
	if (scope === 'cell' && !context.tableCell) throw new Error('A cell must be active (cursor inside) to set padding for only that cell.');
	if (typeof padding !== 'number' || padding < 0) throw new Error(`Invalid padding value: ${padding}.`);
	if (!['cell', 'table'].includes(scope)) throw new Error(`Invalid scope for padding: ${scope}.`);

	let cellsUpdated = 0;
	const applyPadding = (cell) => {
		cell.setPaddingTop(padding);
		cell.setPaddingBottom(padding);
		cell.setPaddingLeft(padding);
		cell.setPaddingRight(padding);
	};

	if (scope === 'cell') {
		applyPadding(context.tableCell);
		cellsUpdated = 1;
	} else { // scope === 'table'
		const numRows = context.table.getNumRows();
		for (let i = 0; i < numRows; i++) {
			const row = context.table.getRow(i);
			const numCells = row.getNumCells();
			for (let j = 0; j < numCells; j++) {
				applyPadding(row.getCell(j));
				cellsUpdated++;
			}
		}
	}
	const scopeText = scope === 'cell' ? 'selected cell' : `all ${cellsUpdated} cells in the table`;
	return {
		success: true,
		message: `Padding set to ${padding}pt for ${scopeText}.`,
		executionTime: new Date().getTime() - startTime
	};
}

/**
 * Sets borders for a table. Individual cell border styling is not directly supported.
 * @param {Object} context - Table context.
 * @param {string} scope - Must be 'table'. 'cell' scope is not supported for distinct borders.
 * @param {number} borderWidth - Border width in points. Use 0 to remove borders.
 * @param {string} [borderColor] - Border color hex string. Defaults to black ('#000000') if borderWidth > 0.
 * @param {number} startTime - Start time for execution tracking.
 * @returns {Object} Result object.
 */
function setBorders(context, scope, borderWidth, borderColor, startTime) {
	if (!context || !context.table) throw new Error('Invalid table context for setBorders.');
	if (scope !== 'table') throw new Error("Border styling is applied to the whole table. 'cell' scope for distinct borders is not supported.");
	if (typeof borderWidth !== 'number' || borderWidth < 0) throw new Error(`Invalid border width: ${borderWidth}.`);

	const effectiveBorderColor = (borderWidth > 0) ? (borderColor || '#000000') : null;

	context.table.setBorderWidth(borderWidth);
	if (borderWidth > 0 && effectiveBorderColor) {
		context.table.setBorderColor(effectiveBorderColor);
	} else if (borderWidth === 0) {
		// When removing borders, explicitly setting color to null ensures it doesn't retain old color
		// though setBorderWidth(0) should be sufficient.
		context.table.setBorderColor(null);
	}

	let message = `Table borders ${borderWidth > 0 ? `set to ${borderWidth}pt` : 'removed'}.`;
	if (borderWidth > 0 && effectiveBorderColor) {
		message += ` Color: ${effectiveBorderColor}.`;
	}

	return {
		success: true,
		message: message,
		executionTime: new Date().getTime() - startTime
	};
}

/**
 * Selects the entire table containing the cursor or the currently selected table.
 * @param {Object} context - Table context.
 * @param {number} startTime - Start time for execution tracking.
 * @returns {Object} Result object.
 */
function selectWholeTable(context, startTime) {
	if (!context || !context.table) throw new Error('Could not identify the table to select.');

	const doc = context.doc;
	doc.setSelection(doc.newRange().addElement(context.table).build());

	return {
		success: true,
		message: 'Table selected.', // This message won't show via UI alert by default
		executionTime: new Date().getTime() - startTime
	};
}