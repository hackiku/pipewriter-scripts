// ai/loadPrompts.js - Handle prompt operations

/**
 * Drop a prompt at the current cursor position in the document
 * @param {Object} params - Parameters object
 * @param {string} params.promptContent - The prompt content to insert
 * @param {string} [params.promptTitle] - Optional title for the prompt
 * @returns {Object} Result object with success/error info
 */
function dropPrompt(params = {}) {
	const startTime = new Date().getTime();

	try {
		if (!params.promptContent) {
			throw new Error('No prompt content provided');
		}

		const doc = DocumentApp.getActiveDocument();
		const cursor = doc.getCursor();
		const selection = doc.getSelection();

		let insertPosition = null;
		let targetElement = null;

		// Determine where to insert the prompt
		if (cursor) {
			// Use cursor position
			targetElement = cursor.getElement();
			const parent = targetElement.getParent();

			// If we're in a paragraph, insert after it
			if (parent.getType() === DocumentApp.ElementType.PARAGRAPH) {
				const body = doc.getBody();
				const elementIndex = body.getChildIndex(parent);
				insertPosition = { container: body, index: elementIndex + 1 };
			} else {
				// Insert at cursor position within the element
				const body = doc.getBody();
				insertPosition = { container: body, index: body.getNumChildren() };
			}
		} else if (selection) {
			// Use selection end
			const rangeElements = selection.getRangeElements();
			if (rangeElements.length > 0) {
				const lastElement = rangeElements[rangeElements.length - 1].getElement();
				const parent = lastElement.getParent();

				if (parent.getType() === DocumentApp.ElementType.PARAGRAPH) {
					const body = doc.getBody();
					const elementIndex = body.getChildIndex(parent);
					insertPosition = { container: body, index: elementIndex + 1 };
				} else {
					const body = doc.getBody();
					insertPosition = { container: body, index: body.getNumChildren() };
				}
			}
		} else {
			// No cursor or selection, insert at end of document
			const body = doc.getBody();
			insertPosition = { container: body, index: body.getNumChildren() };
		}

		if (!insertPosition) {
			throw new Error('Could not determine insertion point');
		}

		// Add some spacing before the prompt
		insertPosition.container.insertParagraph(insertPosition.index, '');
		insertPosition.index++;

		// Add title if provided
		if (params.promptTitle) {
			const titleParagraph = insertPosition.container.insertParagraph(insertPosition.index, params.promptTitle);
			titleParagraph.setHeading(DocumentApp.ParagraphHeading.HEADING3);
			insertPosition.index++;

			// Add a separator line
			insertPosition.container.insertParagraph(insertPosition.index, '---');
			insertPosition.index++;
		}

		// Split prompt content into paragraphs and insert each one
		const promptLines = params.promptContent.split('\n');
		let insertedParagraphs = 0;

		promptLines.forEach((line, lineIndex) => {
			const trimmedLine = line.trim();

			// Insert the line (even if empty, to preserve formatting)
			const newParagraph = insertPosition.container.insertParagraph(insertPosition.index + lineIndex, trimmedLine);

			// Set as normal paragraph
			newParagraph.setHeading(DocumentApp.ParagraphHeading.NORMAL);

			insertedParagraphs++;
		});

		// Add spacing after the prompt
		insertPosition.container.insertParagraph(insertPosition.index + insertedParagraphs, '');

		// Position cursor after the inserted content
		try {
			const newCursorPosition = insertPosition.index + insertedParagraphs + 1;
			if (newCursorPosition < insertPosition.container.getNumChildren()) {
				const nextElement = insertPosition.container.getChild(newCursorPosition);
				doc.setCursor(doc.newPosition(nextElement, 0));
			} else {
				// Add a new paragraph at the end and position cursor there
				const newPara = insertPosition.container.appendParagraph('');
				doc.setCursor(doc.newPosition(newPara, 0));
			}
		} catch (cursorError) {
			Logger.log('Warning: Could not position cursor after prompt insertion: ' + cursorError);
			// Continue anyway, the prompt was still inserted successfully
		}

		return {
			success: true,
			message: 'Prompt inserted successfully',
			insertedParagraphs: insertedParagraphs,
			executionTime: new Date().getTime() - startTime
		};

	} catch (error) {
		Logger.log('Error in dropPrompt: ' + error);
		return {
			success: false,
			error: error.toString(),
			executionTime: new Date().getTime() - startTime
		};
	}
}

/**
 * Menu wrapper for dropping a prompt (for testing from menu)
 */
function menuDropPrompt() {
	return dropPrompt({
		promptContent: "This is a test prompt.\n\nIt has multiple lines to demonstrate the functionality.",
		promptTitle: "Test Prompt"
	});
}