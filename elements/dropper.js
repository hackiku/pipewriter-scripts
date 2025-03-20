// elements/dropper.js - Element dropper functionality

// Attach to global dropper object from Code.js
var dropper = dropper || {};

// #1 data model for UI elements
function createUIElement() {
	return {
		textContent: '',
		formattedContent: '',
		wordCount: 0,
		rowCount: 0,
		cellCount: 0
	};
}

var uiElementIds = [
	'styleguide', 'container-center', 'background-empty', 'background-light', 'container-center-dark', 'background-empty-dark',
	'background-light-dark', 'hero', 'zz-left', 'zz-right', 'zz-left-dark', 'zz-right-dark', 'hero-dark',
	'blurbs-3', 'blurbs-4', 'blurbs-vert-3', 'blurbs-3-dark', 'blurbs-vert-3-dark', 'list-1',
	'list-2', 'list-3', 'list-1-dark', 'list-2-dark', 'list-3-dark', 'button-left', 'button-center',
	'button-right', 'button-2-left', 'button-2-center', 'button-2-right', 'button-left-dark',
	'button-center-dark', 'button-right-dark', 'button-2-left-dark', 'button-2-center-dark', 'button-2-right-dark',
	'cards-2-left', 'cards-2-center', 'cards-3', 'cards-2x2', 'pricing-cards', 'cards-6', 'cards-2-left-dark',
	'cards-2-center-dark', 'cards-3-dark', 'cards-2x2-dark', 'pricing-cards-dark', 'cards-6-dark',
];

var uiElements = {};

uiElementIds.forEach(function (id) {
	uiElements[id] = createUIElement();
});

// #2 get all UI element IDs from master
function populateUIElements() {
	// https://docs.google.com/document/d/1X-mEWo2wuRcVZdA8Y94cFMpUO6tKm8GLxY3ZA8lyulk/edit?usp=sharing
	var masterDocId = '1X-mEWo2wuRcVZdA8Y94cFMpUO6tKm8GLxY3ZA8lyulk';
	// ivan.karaman
	//  var masterDocId = '1uMdieQCJeBQCvkHs7w9ZiVeEB2_cglkF7ZLgeqvxL0U';

	var masterBody = DocumentApp.openById(masterDocId).getBody();

	// Get the total number of child elements in the master document's body.
	var numElements = masterBody.getNumChildren();

	var currentId = null;

	// Iterate over all the child elements in the master document's body.
	for (var i = 0; i < numElements; i++) {
		var element = masterBody.getChild(i);

		// If the element is a paragraph, check if its text is a UI element ID.
		if (element.getType() == DocumentApp.ElementType.PARAGRAPH) {
			var text = element.getText();
			if (uiElements.hasOwnProperty(text)) {
				currentId = text; // Remember the ID to associate the next table with it.
			}
		}
		// If the element is a table and there's a current ID remembered, add the table to the corresponding UI element in uiElements.
		else if (element.getType() == DocumentApp.ElementType.TABLE && currentId !== null) {
			var table = element.copy();
			uiElements[currentId].table = table; // Save a copy of the table.

			// numeric data about table content
			uiElements[currentId].wordCount = table.getText().split(' ').length;
			uiElements[currentId].rowCount = table.getNumRows();
			uiElements[currentId].cellCount = table.getRow(0).getNumCells();
			// TODO add more properties

			// save text content in the object
			var textContent = "";
			for (var r = 0; r < table.getNumRows(); r++) {
				var row = table.getRow(r);
				for (var c = 0; c < row.getNumCells(); c++) {
					var cell = row.getCell(c);
					textContent += cell.getText() + " "; // Add a space between texts from different cells.
				}
			}
			uiElements[currentId].textContent = textContent.trim(); // Remove the trailing space and save the text content.

			currentId = null; // Reset the current ID.
		}
	}
}

// #3 wireframe with elements to cursor
dropper.getElement = function (uiElementId) {
	// if uiElementId is not found in uiElements, return with an error message
	if (!uiElements[uiElementId]) {
		console.log("Invalid UI element requested: " + uiElementId);
		return false;
	}

	// Retrieve the cursor from the active document
	var cursor = DocumentApp.getActiveDocument().getCursor();

	// if no cursor is present, return with an alert message
	if (!cursor) {
		DocumentApp.getUi().alert('Make sure the cursor is blinking');
		return false;
	}

	var cursorElem = cursor.getElement();
	var parent = cursorElem.getParent();
	var body = DocumentApp.getActiveDocument().getBody();

	try {
		// Get the child index of the element where the cursor is.
		var offset = parent.getChildIndex(cursorElem);
	} catch (e) {
		console.log('Error: ', e);
		// If an error occurs (e.g., cursor not inside any child), set offset to 0.
		var offset = 0;
	}

	// Copy the table from the corresponding UI element in uiElements.
	var masterTable = uiElements[uiElementId].table.copy();

	// Check if the cursor is inside a table cell.
	if (parent.getType() == DocumentApp.ElementType.TABLE_CELL) {
		parent.insertTable(offset + 1, masterTable);
	} else {
		body.insertTable(offset + 1, masterTable);
	}

	// Return the UI element info as a simple, JSON-compatible object
	return {
		textContent: cursorElem.asText().getText(),
		wordCount: cursorElem.asText().getText().split(' ').length,
	}
};

// The critical initialization hack: Try to populate UI elements at load time
try {
	populateUIElements();
	for (var id in uiElements) {
		console.log("plain text for '" + id + "': " + uiElements[id].textContent);
	}
} catch (e) {
	console.error("Error during initialization:", e);
}