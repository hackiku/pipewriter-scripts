// tableOps.gs

/**
 * Handles table-related operations in Google Docs
 * @param {Object} params - Parameters passed from the frontend
 * @param {string} params.action - The action to perform
 * @param {Object} params.payload - The payload containing alignment values
 */
function tableOps(params) {
	try {
		const doc = DocumentApp.getActiveDocument();
		const cursor = doc.getCursor();

		if (!cursor) {
			throw new Error('No cursor found in document');
		}

		// Get the current element and find the table cell
		let element = cursor.getElement();
		if (!element) {
			throw new Error('No element found at cursor position');
		}

		// Traverse up the element tree to find the table cell
		while (element && element.getType() !== DocumentApp.ElementType.TABLE_CELL) {
			element = element.getParent();
			if (!element) {
				throw new Error('Cursor is not inside a table cell');
			}
		}

		const tableCell = element;

		switch (params.action) {
			case 'tableAlignHorizontal':
				const hAlignmentMap = {
					'left': DocumentApp.HorizontalAlignment.LEFT,
					'center': DocumentApp.HorizontalAlignment.CENTER,
					'right': DocumentApp.HorizontalAlignment.RIGHT
				};

				if (!hAlignmentMap[params.payload.alignment]) {
					throw new Error('Invalid horizontal alignment value');
				}

				tableCell.setHorizontalAlignment(hAlignmentMap[params.payload.alignment]);
				break;

			case 'tableAlignVertical':
				const vAlignmentMap = {
					'top': DocumentApp.VerticalAlignment.TOP,
					'middle': DocumentApp.VerticalAlignment.CENTER,
					'bottom': DocumentApp.VerticalAlignment.BOTTOM
				};

				if (!vAlignmentMap[params.payload.alignment]) {
					throw new Error('Invalid vertical alignment value');
				}

				tableCell.setVerticalAlignment(vAlignmentMap[params.payload.alignment]);
				break;

			default:
				throw new Error('Unknown table operation');
		}

		return { success: true };

	} catch (error) {
		console.error('Error in tableOps:', error);
		return {
			success: false,
			error: error.toString(),
			message: 'Please ensure your cursor is inside a table cell'
		};
	}
}

/**
 * Test function to verify table operations
 */
function testTable() {
	const testParams = {
		action: 'tableAlignHorizontal',
		payload: {
			alignment: 'center'
		}
	};

	return tableOps(testParams);
}