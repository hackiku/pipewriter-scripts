// appHandler.js - Main handler for app requests (Production)

/**
 * Central handler for all app requests
 * Dispatches to appropriate modules and provides unified error handling
 * No UI alerts - all feedback handled by the client application
 */
var appHandler = (function () {
	// Request start time map (for performance tracking)
	const requestTimes = new Map();

	// Function map for request routing
	const functionMap = {
		// Colors
		'changeBg': function (payload, callback) {
			try {
				const color = payload.color;
				if (!color) {
					throw new Error('No color specified');
				}

				const result = changeBg(color);
				callback(null, result);
			} catch (error) {
				handleError(error, callback);
			}
		},

		'getCurrentColor': function (payload, callback) {
			try {
				const result = getCurrentColor();
				callback(null, result);
			} catch (error) {
				handleError(error, callback);
			}
		},

		// Elements
		'getElement': function (payload, callback) {
			try {
				if (!payload.elementId) {
					throw new Error('No elementId specified');
				}

				// Use indexDropper for faster element insertion if available
				const useIndexDropper = true;

				let result;
				if (useIndexDropper && indexDropper && typeof indexDropper.insertElement === 'function') {
					result = indexDropper.insertElement({
						elementId: payload.elementId,
						theme: payload.theme || 'light'
					});
				} else {
					// Fall back to original dropper
					result = dropper.getElement({
						elementId: payload.elementId,
						theme: payload.theme || 'light'
					});
				}

				callback(null, result);
			} catch (error) {
				handleError(error, callback);
			}
		},

		// HTML operations
		'dropHtml': function (payload, callback) {
			try {
				const result = dropHtml({
					position: payload.position || 'end',
					copyToClipboard: payload.copyToClipboard || false,
					prompt: payload.prompt
				});
				callback(null, result);
			} catch (error) {
				handleError(error, callback);
			}
		},

		'stripHtml': function (payload, callback) {
			try {
				const result = stripHtml({
					all: payload.all || false,
					copyToClipboard: payload.copyToClipboard || false
				});
				callback(null, result);
			} catch (error) {
				handleError(error, callback);
			}
		},

		// Text operations
		'textOps': function (payload, callback) {
			try {
				if (!payload || !payload.action) {
					throw new Error('No text action specified in payload');
				}

				// Call textOps directly with the payload - no UI alerts in production
				const result = textOps(payload);
				callback(null, result);
			} catch (error) {
				handleError(error, callback);
			}
		},

		// Style operations
		'getStyleTemplate': function (payload, callback) {
			try {
				const result = getStyleTemplate({
					templateId: payload.templateId || 'style-minimal',
					theme: payload.theme || 'light'
				});
				callback(null, result);
			} catch (error) {
				handleError(error, callback);
			}
		},

		// Table operations
		'tableOps': function (payload, callback) {
			try {
				if (!payload || !payload.action) {
					throw new Error('No table action specified in payload');
				}

				// Call tableOps directly with the payload - no UI alerts in production
				const result = tableOps(payload);
				callback(null, result);
			} catch (error) {
				handleError(error, callback);
			}
		}
	};

	/**
	 * Centralized error handler
	 * @param {Error} error - The error to handle
	 * @param {Function} callback - Callback function
	 */
	function handleError(error, callback) {
		Logger.log('Error in request handler: ' + error);
		callback(error.toString(), null);
	}

	/**
	 * Process a request with performance tracking
	 * @param {string} functionName - The function to call
	 * @param {Object} payload - The payload for the function
	 * @param {Function} callback - Callback function
	 */
	function processRequest(functionName, payload, callback) {
		const requestId = `${functionName}-${Date.now()}`;
		requestTimes.set(requestId, Date.now());

		// Create a wrapped callback that adds execution time
		const wrappedCallback = function (error, result) {
			const startTime = requestTimes.get(requestId);
			const executionTime = startTime ? Date.now() - startTime : null;

			// Clean up request time entry
			requestTimes.delete(requestId);

			// Add execution time to successful results
			if (!error && result) {
				result.executionTime = executionTime;
			}

			// Call the original callback
			callback(error, result);
		};

		// Call the appropriate function
		if (functionMap[functionName]) {
			functionMap[functionName](payload, wrappedCallback);
		} else {
			handleError(new Error(`Unknown function: ${functionName}`), callback);
		}
	}

	return {
		processRequest: processRequest
	};
})();

/**
 * Global function to handle all app requests
 * @param {string} functionName - The function to call
 * @param {Object} payload - The function parameters
 * @returns {Object} - Function result
 */
function handleAppRequest(functionName, payload = {}) {
	let result = null;

	appHandler.processRequest(functionName, payload, function (error, response) {
		if (error) {
			result = {
				success: false,
				error: error,
				functionName: functionName
			};
		} else {
			result = {
				success: true,
				...response,
				functionName: functionName
			};
		}
	});

	return result;
}