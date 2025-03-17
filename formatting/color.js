// colorOps.gs

/**
 * Changes the background color of the active document
 * @param {string} color - The hex color code
 * @returns {Object} - Result of the operation
 */

function changeBg(color) {
	const startTime = new Date().getTime();
	try {
		const body = DocumentApp.getActiveDocument().getBody();
		body.setBackgroundColor(color);

		return {
			success: true,
			executionTime: new Date().getTime() - startTime
		};
	} catch (error) {
		Logger.log('Error in changeBg:', error);
		return {
			success: false,
			error: error.toString(),
			executionTime: new Date().getTime() - startTime
		};
	}
}

/**
 * Gets the current background color of the document
 * @returns {Object} - Result of the operation with the current color
 */
function getCurrentColor() {
	const startTime = new Date().getTime();

	try {
		const body = DocumentApp.getActiveDocument().getBody();
		const color = body.getBackgroundColor();

		return {
			success: true,
			color: color || '#FFFFFF', // Default to white if no color is set
			executionTime: new Date().getTime() - startTime
		};
	} catch (error) {
		Logger.log('Error in getCurrentColor:', error);
		return {
			success: false,
			error: error.toString(),
			executionTime: new Date().getTime() - startTime
		};
	}
}