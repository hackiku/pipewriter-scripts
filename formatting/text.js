// formatting/text.js - Text operations for wireframing app

/**
 * Main text operations function
 * @param {Object} params - Operation parameters
 * @param {string} params.action - The action to perform ('applyStyle', 'updateAllMatching', 'getStyleInfo', 'getAllStyles')
 * @param {string} [params.headingType] - Heading type for applyStyle ('NORMAL', 'HEADING1'-'HEADING6')
 * @returns {Object} Result object with success/error info
 */
function textOps(params) {
	const startTime = new Date().getTime();

	try {
		if (!params || !params.action) {
			throw new Error('No action specified for textOps');
		}

		const context = getTextContext();
		if (!context.success) {
			throw new Error(context.error);
		}

		let result;
		switch (params.action) {
			case 'applyStyle':
				result = applyTextStyle(context, params.headingType, startTime);
				break;

			case 'updateAllMatching':
				result = updateAllMatchingHeadings(context, startTime);
				break;

			case 'getStyleInfo':
				// For app use - returns data without UI alert
				result = getStyleInfoForApp(context, startTime);
				break;

			case 'getAllStyles':
				// NEW: Get all styles used in the document
				result = getAllDocumentStylesForApp(context, startTime);
				break;

			default:
				throw new Error(`Unknown text action: ${params.action}`);
		}

		return result;

	} catch (error) {
		const errorResult = {
			success: false,
			error: error.toString(),
			message: error.message || 'Please place cursor in document.',
			executionTime: new Date().getTime() - startTime
		};

		Logger.log(`Error in textOps (action: ${params.action}): ${error.message}`);
		return errorResult;
	}
}

/**
 * Get text context (paragraph, cursor, selection info)
 * @returns {Object} Context object or error
 */
function getTextContext() {
	try {
		const doc = DocumentApp.getActiveDocument();
		const cursor = doc.getCursor();
		const selection = doc.getSelection();

		let paragraph = null;
		let textElement = null;
		let contextType = 'none';

		// Try cursor first
		if (cursor) {
			const element = cursor.getElement();
			if (element.getType() === DocumentApp.ElementType.TEXT) {
				paragraph = element.getParent().asParagraph();
				textElement = element.asText();
				contextType = 'cursor';
			} else if (element.getType() === DocumentApp.ElementType.PARAGRAPH) {
				paragraph = element.asParagraph();
				// Get first text element in paragraph for styling
				if (paragraph.getNumChildren() > 0) {
					const firstChild = paragraph.getChild(0);
					if (firstChild.getType() === DocumentApp.ElementType.TEXT) {
						textElement = firstChild.asText();
					}
				}
				contextType = 'cursor';
			}
		}

		// Fallback to selection
		if (!paragraph && selection) {
			const rangeElements = selection.getRangeElements();
			if (rangeElements.length > 0) {
				const element = rangeElements[0].getElement();
				if (element.getType() === DocumentApp.ElementType.TEXT) {
					paragraph = element.getParent().asParagraph();
					textElement = element.asText();
					contextType = 'selection';
				} else if (element.getType() === DocumentApp.ElementType.PARAGRAPH) {
					paragraph = element.asParagraph();
					if (paragraph.getNumChildren() > 0) {
						const firstChild = paragraph.getChild(0);
						if (firstChild.getType() === DocumentApp.ElementType.TEXT) {
							textElement = firstChild.asText();
						}
					}
					contextType = 'selection';
				}
			}
		}

		if (!paragraph) {
			return {
				success: false,
				error: 'No cursor or selection found. Please place cursor in text.'
			};
		}

		return {
			success: true,
			paragraph: paragraph,
			textElement: textElement,
			cursor: cursor,
			selection: selection,
			contextType: contextType,
			doc: doc
		};

	} catch (error) {
		return {
			success: false,
			error: 'Failed to get text context: ' + error.message
		};
	}
}

/**
 * Apply text style to current paragraph
 */
function applyTextStyle(context, headingType, startTime) {
	try {
		const headingMap = {
			'NORMAL': DocumentApp.ParagraphHeading.NORMAL,
			'HEADING1': DocumentApp.ParagraphHeading.HEADING1,
			'HEADING2': DocumentApp.ParagraphHeading.HEADING2,
			'HEADING3': DocumentApp.ParagraphHeading.HEADING3,
			'HEADING4': DocumentApp.ParagraphHeading.HEADING4,
			'HEADING5': DocumentApp.ParagraphHeading.HEADING5,
			'HEADING6': DocumentApp.ParagraphHeading.HEADING6
		};

		const heading = headingMap[headingType];
		if (heading === undefined) {
			throw new Error(`Invalid heading type: ${headingType}`);
		}

		context.paragraph.setHeading(heading);

		return {
			success: true,
			message: `Applied ${getHeadingDisplayName(heading)} style`,
			headingType: headingType,
			executionTime: new Date().getTime() - startTime
		};

	} catch (error) {
		throw new Error(`Failed to apply text style: ${error.message}`);
	}
}

/**
 * Update all matching headings to match current paragraph style
 */
function updateAllMatchingHeadings(context, startTime) {
	try {
		const sourceParagraph = context.paragraph;
		const sourceHeading = sourceParagraph.getHeading();

		if (sourceHeading === DocumentApp.ParagraphHeading.NORMAL) {
			throw new Error('Current paragraph is Normal text. Please select a heading to use as template.');
		}

		// Get ALL formatting from the source paragraph's text
		let sourceTextAttributes = {};
		let sourceParagraphAttributes = {};

		// Get text attributes from the first text element
		if (context.textElement) {
			sourceTextAttributes = getAllTextAttributes(context.textElement);
		}

		// Get paragraph-level attributes
		sourceParagraphAttributes = cleanNullAttributes(sourceParagraph.getAttributes());

		// Find all matching headings
		const body = context.doc.getBody();
		const allParagraphs = body.getParagraphs();
		const matchingParagraphs = allParagraphs.filter(p =>
			p.getHeading() === sourceHeading && p !== sourceParagraph
		);

		// Apply updates to each matching paragraph
		let updated = 0;
		matchingParagraphs.forEach(paragraph => {
			// Apply paragraph-level attributes
			paragraph.setAttributes(sourceParagraphAttributes);

			// Apply text-level attributes to all text in the paragraph
			if (Object.keys(sourceTextAttributes).length > 0) {
				const text = paragraph.editAsText();
				if (text.getText().length > 0) {
					text.setAttributes(0, text.getText().length - 1, sourceTextAttributes);
				}
			}
			updated++;
		});

		// Update document-wide heading style definition
		const combinedAttributes = { ...sourceParagraphAttributes, ...sourceTextAttributes };
		body.setHeadingAttributes(sourceHeading, combinedAttributes);

		const headingName = getHeadingDisplayName(sourceHeading);

		return {
			success: true,
			message: `Updated ${updated} matching ${headingName} paragraphs to match selected style`,
			updatedCount: updated,
			executionTime: new Date().getTime() - startTime
		};

	} catch (error) {
		throw new Error(`Failed to update matching headings: ${error.message}`);
	}
}

/**
 * Get current paragraph style information FOR APP USE (no UI alert)
 * Returns proper data structure for Svelte app with PROPER STYLE EXTRACTION
 */
function getStyleInfoForApp(context, startTime) {
	try {
		const paragraph = context.paragraph;
		const heading = paragraph.getHeading();
		const text = paragraph.getText().substring(0, 50) + (paragraph.getText().length > 50 ? '...' : '');
		const body = context.doc.getBody();

		// ENHANCED: Get the actual applied formatting by combining document-level and text-level attributes
		let textAttributes = getEffectiveTextAttributes(paragraph, context.textElement, body);
		const paragraphAttributes = cleanNullAttributes(paragraph.getAttributes());
		const headingName = getHeadingDisplayName(heading);

		// Count matching paragraphs
		const allParagraphs = body.getParagraphs();
		const matchingCount = allParagraphs.filter(p => p.getHeading() === heading).length;

		// Return structured data for app consumption with proper attributes
		return {
			success: true,
			message: 'Style info retrieved',
			data: {
				textAttributes: textAttributes,
				paragraphAttributes: paragraphAttributes,
				heading: heading, // Return the actual heading enum for proper mapping
				headingName: headingName,
				text: text,
				matchingCount: matchingCount
			},
			executionTime: new Date().getTime() - startTime
		};

	} catch (error) {
		throw new Error(`Failed to get style info: ${error.message}`);
	}
}

/**
 * NEW: Get all styles used in the document FOR APP USE
 * Scans all paragraphs and returns unique heading styles with their attributes
 */
function getAllDocumentStylesForApp(context, startTime) {
	try {
		const body = context.doc.getBody();
		const allParagraphs = body.getParagraphs();
		const uniqueStyles = new Map();

		// Scan all paragraphs to find unique heading types
		allParagraphs.forEach((paragraph, index) => {
			const heading = paragraph.getHeading();
			const headingKey = heading.toString();

			// Skip if we've already processed this heading type
			if (uniqueStyles.has(headingKey)) {
				return;
			}

			// Get the first text element for this paragraph type
			let textElement = null;
			if (paragraph.getNumChildren() > 0) {
				const firstChild = paragraph.getChild(0);
				if (firstChild.getType() === DocumentApp.ElementType.TEXT) {
					textElement = firstChild.asText();
				}
			}

			// Get effective text attributes for this heading type
			const textAttributes = getEffectiveTextAttributes(paragraph, textElement, body);
			const paragraphAttributes = cleanNullAttributes(paragraph.getAttributes());
			const headingName = getHeadingDisplayName(heading);
			const sampleText = paragraph.getText().substring(0, 30) + (paragraph.getText().length > 30 ? '...' : '');

			// Count all paragraphs of this heading type
			const matchingCount = allParagraphs.filter(p => p.getHeading() === heading).length;

			uniqueStyles.set(headingKey, {
				textAttributes: textAttributes,
				paragraphAttributes: paragraphAttributes,
				heading: heading,
				headingName: headingName,
				sampleText: sampleText,
				matchingCount: matchingCount
			});
		});

		// Convert Map to Array for return
		const stylesArray = Array.from(uniqueStyles.values());

		return {
			success: true,
			message: `Found ${stylesArray.length} unique text styles in document`,
			data: {
				styles: stylesArray,
				totalParagraphs: allParagraphs.length
			},
			executionTime: new Date().getTime() - startTime
		};

	} catch (error) {
		throw new Error(`Failed to get all document styles: ${error.message}`);
	}
}

/**
 * NEW: Get effective text attributes by combining document-level heading styles with text-level overrides
 * This solves the "undefined" attribute problem by getting the actual applied formatting
 */
function getEffectiveTextAttributes(paragraph, textElement, body) {
	const heading = paragraph.getHeading();
	let effectiveAttributes = {};

	try {
		// Step 1: Get document-level heading attributes (the base style)
		const documentHeadingAttributes = body.getHeadingAttributes(heading);
		if (documentHeadingAttributes) {
			effectiveAttributes = { ...cleanNullAttributes(documentHeadingAttributes) };
		}

		// Step 2: Get text-level attributes (overrides) if we have a text element
		if (textElement) {
			const textLevelAttributes = getAllTextAttributes(textElement);
			// Merge text-level attributes over document-level (text-level takes precedence)
			effectiveAttributes = { ...effectiveAttributes, ...textLevelAttributes };
		}

		// Step 3: Clean up and format for consistent API response
		const cleanedAttributes = {};

		// Map DocumentApp.Attribute keys to string keys for consistent API
		Object.keys(effectiveAttributes).forEach(key => {
			const value = effectiveAttributes[key];
			if (value !== null && value !== undefined) {
				// Convert attribute enum to string if needed
				const stringKey = key.toString();
				cleanedAttributes[stringKey] = value;
			}
		});

		return cleanedAttributes;

	} catch (error) {
		Logger.log('Error getting effective text attributes: ' + error);
		// Fallback to just text-level attributes if document-level fails
		return textElement ? getAllTextAttributes(textElement) : {};
	}
}

/**
 * Get current paragraph style information WITH UI ALERT (for menu use)
 * Shows the alert like the working bound script
 */
function getStyleInfoWithAlert(context, startTime) {
	try {
		const paragraph = context.paragraph;
		const heading = paragraph.getHeading();
		const text = paragraph.getText().substring(0, 50) + (paragraph.getText().length > 50 ? '...' : '');
		const body = context.doc.getBody();

		// ENHANCED: Use the new effective attributes function
		const textAttributes = getEffectiveTextAttributes(paragraph, context.textElement, body);
		const paragraphAttributes = paragraph.getAttributes();
		const headingName = getHeadingDisplayName(heading);

		// Count matching paragraphs
		const allParagraphs = body.getParagraphs();
		const matchingCount = allParagraphs.filter(p => p.getHeading() === heading).length;

		// Build and show the alert with ACTUAL formatting info
		const info = [
			`Text: "${text}"`,
			`Style: ${headingName}`,
			`Total ${headingName} paragraphs: ${matchingCount}`,
			``,
			`Text Formatting:`,
			`Font: ${textAttributes[DocumentApp.Attribute.FONT_FAMILY] || 'Default'}`,
			`Size: ${textAttributes[DocumentApp.Attribute.FONT_SIZE] || 'Default'}`,
			`Bold: ${textAttributes[DocumentApp.Attribute.BOLD] ? 'Yes' : 'No'}`,
			`Italic: ${textAttributes[DocumentApp.Attribute.ITALIC] ? 'Yes' : 'No'}`,
			`Color: ${textAttributes[DocumentApp.Attribute.FOREGROUND_COLOR] || 'Default'}`,
			``,
			`Paragraph Formatting:`,
			`Alignment: ${paragraphAttributes[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] || 'Default'}`,
			`Line Spacing: ${paragraphAttributes[DocumentApp.Attribute.LINE_SPACING] || 'Default'}`
		].join('\n');

		DocumentApp.getUi().alert('Current Paragraph Style Info:\n\n' + info);

		return {
			success: true,
			message: 'Style info displayed',
			textAttributes: textAttributes,
			paragraphAttributes: paragraphAttributes,
			executionTime: new Date().getTime() - startTime
		};

	} catch (error) {
		throw new Error(`Failed to get style info: ${error.message}`);
	}
}

/**
 * Get all text attributes from text element (improved from working script)
 */
function getAllTextAttributes(textElement) {
	const attributes = {};

	try {
		// Get all possible text attributes
		const textAttrs = textElement.getAttributes();

		// Include all non-null attributes with better null checking
		const textAttributeKeys = [
			DocumentApp.Attribute.FONT_FAMILY,
			DocumentApp.Attribute.FONT_SIZE,
			DocumentApp.Attribute.BOLD,
			DocumentApp.Attribute.ITALIC,
			DocumentApp.Attribute.UNDERLINE,
			DocumentApp.Attribute.STRIKETHROUGH,
			DocumentApp.Attribute.FOREGROUND_COLOR,
			DocumentApp.Attribute.BACKGROUND_COLOR,
			DocumentApp.Attribute.LINK_URL
		];

		textAttributeKeys.forEach(attr => {
			try {
				const value = textAttrs[attr];
				// Only include non-null, non-undefined values
				if (value !== null && value !== undefined) {
					attributes[attr] = value;
				}
			} catch (attrError) {
				// Skip attributes that can't be read
				Logger.log(`Could not read attribute ${attr}: ${attrError}`);
			}
		});

	} catch (error) {
		Logger.log('Error getting text attributes: ' + error);
	}

	return attributes;
}

/**
 * Clean null attributes from object
 */
function cleanNullAttributes(attributes) {
	const cleaned = {};
	for (const attr in attributes) {
		if (attributes[attr] !== null && attributes[attr] !== undefined) {
			cleaned[attr] = attributes[attr];
		}
	}
	return cleaned;
}

/**
 * Get display name for heading type
 */
function getHeadingDisplayName(heading) {
	const headingNames = {
		[DocumentApp.ParagraphHeading.NORMAL]: 'Normal text',
		[DocumentApp.ParagraphHeading.HEADING1]: 'Heading 1',
		[DocumentApp.ParagraphHeading.HEADING2]: 'Heading 2',
		[DocumentApp.ParagraphHeading.HEADING3]: 'Heading 3',
		[DocumentApp.ParagraphHeading.HEADING4]: 'Heading 4',
		[DocumentApp.ParagraphHeading.HEADING5]: 'Heading 5',
		[DocumentApp.ParagraphHeading.HEADING6]: 'Heading 6'
	};
	return headingNames[heading] || 'Unknown';
}