// textOps.gs

// Style guide template IDs
const STYLE_TEMPLATES = {
	'style-minimal': { light: 'style-minimal', dark: 'style-minimal-dark' },
	'style-detailed': { light: 'style-detailed', dark: 'style-detailed-dark' }
};

/**
 * Get style guide template from master doc
 * @param {Object} params - Parameters for style template
 * @param {string} params.templateId - Template ID ('style-minimal' or 'style-detailed')
 * @param {string} params.theme - Theme ('light' or 'dark')
 * @returns {Object} Response with success/error and optional table
 */

function getStyleTemplate(params = {}) {
	const startTime = new Date().getTime();

	try {
		const { templateId = 'style-minimal', theme = 'light' } = params;
		Logger.log(`Getting style template: ${templateId} (${theme})`);

		// Validate inputs
		if (!STYLE_TEMPLATES[templateId]) {
			throw new Error(`Invalid template ID: ${templateId}`);
		}

		if (!['light', 'dark'].includes(theme)) {
			throw new Error(`Invalid theme: ${theme}`);
		}

		// Get template ID based on theme
		const targetId = STYLE_TEMPLATES[templateId][theme];

		// Get table from master doc using existing infrastructure 
		const table = getElementFromMaster(targetId, theme);

		if (!table) {
			throw new Error(`Template ${targetId} not found`);
		}

		// Insert at cursor position using existing helper
		const inserted = insertElementTable(table);
		if (!inserted) {
			throw new Error('Failed to insert style guide');
		}

		return {
			success: true,
			templateId: targetId,
			theme,
			executionTime: new Date().getTime() - startTime
		};

	} catch (error) {
		Logger.log('Error in getStyleTemplate:', error);
		return {
			success: false,
			error: error.toString(),
			executionTime: new Date().getTime() - startTime
		};
	}
}

// Testing utilities for IDE development
function testStyleTemplate() {
	// Test light minimal
	Logger.log('\nTesting light minimal:');
	Logger.log(getStyleTemplate({
		templateId: 'style-minimal',
		theme: 'light'
	}));

	// Test dark minimal
	Logger.log('\nTesting dark minimal:');
	Logger.log(getStyleTemplate({
		templateId: 'style-minimal',
		theme: 'dark'
	}));

	// Test light detailed
	Logger.log('\nTesting light detailed:');
	Logger.log(getStyleTemplate({
		templateId: 'style-detailed',
		theme: 'light'
	}));

	// Test error case
	Logger.log('\nTesting invalid template:');
	Logger.log(getStyleTemplate({
		templateId: 'invalid',
		theme: 'light'
	}));
}

// Menu integration helper
function insertStyleGuide() {
	return getStyleTemplate({
		templateId: 'style-minimal',
		theme: 'light'
	});
}