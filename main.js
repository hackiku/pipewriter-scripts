// main.js - Core initialization and menu setup

/**
 * Triggered when the add-on is installed
 */
function onInstall(e) {
	try {
		onOpen(e);
	} catch (e) {
		// Silent fail for authorization during install
		console.error("Install error:", e);
	}
}

/**
 * Triggered when the document is opened
 * Sets up the Pipewriter menu
 */
function onOpen() {
	var ui = DocumentApp.getUi();
	ui.createMenu("Pipewriter")
		.addItem("Open App", "showFormInSidebar")
		.addSeparator()
		// HTML operations
		.addItem("HTML to start ↑", "dropHtmlStart")
		.addItem("HTML to end ↓", "dropHtmlEnd")
		.addItem("Strip HTML tags", "stripHtmlTags")
		.addItem("Delete HTML", "stripHtmlAll")
		.addItem("HTML to clipboard", "dropHtmlClipboard")
		.addToUi();
}

/**
 * Shows the Pipewriter sidebar
 */
function showFormInSidebar() {
	var form = HtmlService.createTemplateFromFile("Index")
		.evaluate()
		.setTitle("Pipewriter");
	DocumentApp.getUi().showSidebar(form);
}

// Wrapper functions for menu items - these call functions in other files
function dropHtmlStart() {
	return dropHtml({ position: 'start' });
}

function dropHtmlEnd() {
	return dropHtml({ position: 'end' });
}

function dropHtmlClipboard() {
	return dropHtml({ copyToClipboard: true });
}

// Initialize necessary components with error handling
try {
	// Any initialization code needed at script load
	// For example, preloading commonly used elements
} catch (e) {
	// Silent fail for authorization during load
	console.error("Initialization error:", e);
}