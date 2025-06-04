// formatting/table.js - Clean, consolidated table operations

/**
 * Main table operations function
 * @param {Object} params - Operation parameters
 * @param {string} params.action - The action to perform
 * @param {string} [params.scope] - Scope for operations ('cell', 'table')
 * @param {string} [params.alignment] - Alignment value
 * @param {number} [params.padding] - Padding value in points
 * @param {number} [params.borderWidth] - Border width in points
 * @param {string} [params.borderColor] - Border color hex string
 * @param {boolean} [params.showAlert=true] - Whether to show UI alerts
 * @returns {Object} Result object with success/error info
 */
function tableOps(params) {
	const startTime = new Date().getTime();
	const showAlert = (params.showAlert === undefined) ? true : params.showAlert;
	let uiAlertMessage = null;

	try {
		if (!params || !params.action) {
			throw new Error('No action specified for tableOps');
		}

		// Get table context - some actions like selectWholeTable work with table selection
		const allowTableSelection = params.action === 'selectWholeTable' || params.action === 'openTableOptions';
		const context = getTableContext(allowTableSelection);

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
				// No UI alert for selection - it's a visual change
				break;

			case 'openTableOptions':
				result = openTableOptions(context, startTime);
				uiAlertMessage = result.message;
				break;

			default:
				throw new Error(`Unknown table action: ${params.action}`);
		}

		// Show success alert if enabled and we have a message
		if (result.success && uiAlertMessage && showAlert) {
			DocumentApp.getUi().alert('✅ Table Control', uiAlertMessage, DocumentApp.getUi().ButtonSet.OK);
		}

		return result;

	} catch (error) {
		const errorResult = {
			success: false,
			error: error.toString(),
			message: error.message || 'Please ensure your cursor is inside a table cell for most operations.',
			executionTime: new Date().getTime() - startTime
		};

		Logger.log(`Error in tableOps (action: ${params.action}): ${error.message}`);

		if (showAlert) {
			DocumentApp.getUi().alert('❌ Table Control Error', errorResult.message, DocumentApp.getUi().ButtonSet.OK);
		}

		return errorResult;
	}
}

/**
 * Get table context (table, cell, cursor info)
 * @param {boolean} allowTableSelectionOnly - If true, allows context even if only a table is selected
 * @returns {Object} Context object or error
 */
function getTableContext(allowTableSelectionOnly = false) {
	try {
		const doc = DocumentApp.getActiveDocument();
		const cursor = doc.getCursor();
		const selection = doc.getSelection();

		let table, tableCell = null, tableRow = null;

		// First try to get table from cursor position
		if (cursor) {
			let element = cursor.getElement();
			if (element) {
				let currentElement = element;
				let attempts = 0;
				const maxAttempts = 10;

				// Traverse up to find table cell
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
		}

		// If no table found via cursor and table selection is allowed, check selection
		if (!table && allowTableSelectionOnly && selection) {
			const rangeElements = selection.getRangeElements();
			if (rangeElements.length === 1 && rangeElements[0].getElement().getType() === DocumentApp.ElementType.TABLE) {
				table = rangeElements[0].getElement().asTable();
			}
		}

		if (!table) {
			return {
				success: false,
				error: 'Cursor is not inside a table cell, nor is a table selected. Please click inside a table or select a table.'
			};
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
		return {
			success: false,
			error: 'Failed to get table context: ' + error.message
		};
	}
}

/**
 * Set cell content vertical alignment
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
 * Set cell padding
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
 * Set table borders
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
 * Select the entire table
 */
function selectWholeTable(context, startTime) {
	if (!context || !context.table) throw new Error('Could not identify the table to select.');

	const doc = context.doc;
	doc.setSelection(doc.newRange().addElement(context.table).build());

	return {
		success: true,
		message: 'Table selected.',
		executionTime: new Date().getTime() - startTime
	};
}

/**
 * Open table options (simulates right-click > Table options)
 * Note: This is a workaround since we can't actually trigger the native menu
 */
function openTableOptions(context, startTime) {
	if (!context || !context.table) throw new Error('Could not identify the table for options.');

	// Since we can't actually open the native table options dialog,
	// we'll select the table and show a helpful message
	const doc = context.doc;
	doc.setSelection(doc.newRange().addElement(context.table).build());

	return {
		success: true,
		message: 'Table selected. Right-click the table to access native table options.',
		executionTime: new Date().getTime() - startTime
	};
}