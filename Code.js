// main.js - Entry point and menu setup

// Expose global functions
var dropper = {};

function onInstall(e) {
	onOpen(e);
}

function onOpen() {
	var ui = DocumentApp.getUi();
	ui.createMenu("Pipewriter")
		.addItem("App", "showFormInSidebar")
		.addSeparator()
		.addSubMenu(ui.createMenu("Index Insert (Fast)")
			.addItem("Zigzag Left", "indexDropper.insertZigzagLeft")
			.addItem("Blurbs 3", "indexDropper.insertBlurbs3")
		)
		.addSeparator()
		.addItem("Insert Zigzag Right", "insertZigzagRight")
		.addSeparator()
		.addItem("HTML to start ↑", "dropHtmlStart")
		.addItem("HTML to end ↓", "dropHtmlEnd")
		.addItem("Strip HTML tags", "stripHtmlTags")
		.addItem("Delete HTML", "stripHtmlAll")
		.addItem("HTML to clipboard", "dropHtmlClipboard")
		.addSeparator()
		.addItem("format to HTML", "doc2html")
		.addItem("delete </>", "deleteHTMLtags")
		.addItem("HTML to Pipe", "html2doc")
		.addItem("Gray Background", "graybg")
		.addToUi();
}

function showFormInSidebar() {
	var form = HtmlService.createTemplateFromFile('Index').evaluate().setTitle('Pipewriter');
	var userProperties = PropertiesService.getUserProperties();
	DocumentApp.getUi().showSidebar(form);
}

// function tableOps(params) {
// 	// Forward to the tableOps function in table.js
// 	return tableOps(params);
// }

// Wrapper functions for menu items
function dropHtmlStart() {
	return dropHtml({ position: 'start' });
}

function dropHtmlEnd() {
	return dropHtml({ position: 'end' });
}

function dropHtmlClipboard() {
	return dropHtml({ copyToClipboard: true });
}

// Element getter wrapper - forwards to dropper.js
function getElement(elementId) {
	return dropper.getElement(elementId);
}

function insertZigzagRight() {
	return tableCreator.createZigzagRight();
}