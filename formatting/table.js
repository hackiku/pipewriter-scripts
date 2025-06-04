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
 * @returns {Object} Result object with success/error info
 */
function tableOps(params) {
	const startTime = new Date().getTime();

	try {
		if (!params || !params.action) {
			throw new Error('No action specified for tableOps');
		}

		// Get table context - selectWholeTable works with table selection
		const allowTableSelection = params.action === 'selectWholeTable';
		const context = getTableContext(allowTableSelection);

		if (!context.success) {
			throw new Error(context.error);
		}

		let result;
		switch (params.action) {
			case 'setCellAlignment':
				result = setCellAlignment(context, params.scope, params.alignment, startTime);
				break;

			case 'setCellPadding':
				result = setCellPadding(context, params.scope, params.padding, startTime);
				break;

			case 'setBorders':
				result = setBorders(context, params.scope, params.borderWidth, params.borderColor, startTime);
				break;

			case 'setCellBackground':
				result = setCellBackground(context, params.scope, params.backgroundColor, startTime);
				break;

			case 'selectWholeTable':
				result = selectWholeTable(context, startTime);
				break;

			default:
				throw new Error(`Unknown table action: ${params.action}`);
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
	if (!['top', 'middle', 'bottom'].includes(alignment)) throw new Error(`Invalid cell alignment: ${alignment}.`);

	const alignmentMap = {
		'top': DocumentApp.VerticalAlignment.TOP,
		'middle': DocumentApp.VerticalAlignment.CENTER,
		'bottom': DocumentApp.VerticalAlignment.BOTTOM
	};

	const verticalAlignment = alignmentMap[alignment];
	const cellOperation = (cell) => cell.setVerticalAlignment(verticalAlignment);
	const { scopeText } = applyCellOperation(context, scope, cellOperation, 'cell alignment');

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
	if (typeof padding !== 'number' || padding < 0) throw new Error(`Invalid padding value: ${padding}.`);

	const cellOperation = (cell) => {
		cell.setPaddingTop(padding);
		cell.setPaddingBottom(padding);
		cell.setPaddingLeft(padding);
		cell.setPaddingRight(padding);
	};

	const { scopeText } = applyCellOperation(context, scope, cellOperation, 'padding');

	return {
		success: true,
		message: `Padding set to ${padding}pt for ${scopeText}.`,
		executionTime: new Date().getTime() - startTime
	};
}

/**
 * Abstract helper for applying operations to current cell vs all cells
 * @param {Object} context - Table context
 * @param {string} scope - 'cell' or 'table'
 * @param {Function} cellOperation - Function to apply to each cell
 * @param {string} operationName - Name of operation for error messages
 * @returns {Object} - { cellsUpdated: number, scopeText: string }
 */
function applyCellOperation(context, scope, cellOperation, operationName) {
	if (!context || !context.table) throw new Error(`Invalid table context for ${operationName}.`);
	if (scope === 'cell' && !context.tableCell) throw new Error(`A cell must be active (cursor inside) to apply ${operationName} to only that cell.`);
	if (!['cell', 'table'].includes(scope)) throw new Error(`Invalid scope for ${operationName}: ${scope}.`);

	let cellsUpdated = 0;

	if (scope === 'cell') {
		cellOperation(context.tableCell);
		cellsUpdated = 1;
	} else { // scope === 'table'
		const numRows = context.table.getNumRows();
		for (let i = 0; i < numRows; i++) {
			const row = context.table.getRow(i);
			const numCells = row.getNumCells();
			for (let j = 0; j < numCells; j++) {
				cellOperation(row.getCell(j));
				cellsUpdated++;
			}
		}
	}

	const scopeText = scope === 'cell' ? 'selected cell' : `all ${cellsUpdated} cells in the table`;
	return { cellsUpdated, scopeText };
}

/**
 * Set table borders (table-wide only)
 */
function setBorders(context, scope, borderWidth, borderColor, startTime) {
	if (!context || !context.table) throw new Error('Invalid table context for setBorders.');
	if (scope !== 'table') throw new Error("Border styling is only supported for the whole table, not individual cells.");
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
 * Set cell background colors
 */
function setCellBackground(context, scope, backgroundColor, startTime) {
	if (!backgroundColor) throw new Error('No background color specified.');

	const cellOperation = (cell) => {
		cell.setBackgroundColor(backgroundColor);
	};

	const { scopeText } = applyCellOperation(context, scope, cellOperation, 'background color');

	return {
		success: true,
		message: `Background color set to ${backgroundColor} for ${scopeText}.`,
		executionTime: new Date().getTime() - startTime
	};
}

/**
 * Select the entire table (precisely, without extra elements)
 */
function selectWholeTable(context, startTime) {
	if (!context || !context.table) throw new Error('Could not identify the table to select.');

	try {
		const doc = context.doc;
		const table = context.table;

		// Create a precise selection of just the table
		const range = doc.newRange();
		range.addElement(table);
		doc.setSelection(range.build());

		return {
			success: true,
			message: 'Table selected.',
			executionTime: new Date().getTime() - startTime
		};
	} catch (error) {
		throw new Error('Failed to select table: ' + error.toString());
	}
}