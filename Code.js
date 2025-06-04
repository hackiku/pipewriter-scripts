// Code.js - Entry point and menu setup

// Expose global functions
var dropper = {};

function onInstall(e) {
	onOpen(e);
}

function onOpen() {
	var ui = DocumentApp.getUi();
	ui.createMenu("Pipewriter")
		.addItem("Open Sidebar", "showFormInSidebar")
		.addSeparator()

		// Quick Element Insertions
		.addSubMenu(ui.createMenu("🎯 Quick Insert")
			.addItem("Hero Section", "menuInsertHero")
			.addItem("Zigzag Left", "menuInsertZigzagLeft")
			.addItem("Zigzag Right", "menuInsertZigzagRight")
			.addItem("Blurbs 3-Column", "menuInsertBlurbs3")
			.addItem("Cards 2-Column", "menuInsertCards2")
			.addItem("Cards 3-Column", "menuInsertCards3")
		)

		.addSeparator()

		// Table Controls
		.addSubMenu(ui.createMenu("📋 Table Controls")
			.addItem("Select Whole Table", "menuSelectTable")
			.addItem("Open Table Options", "menuOpenTableOptions")
			.addSeparator()

			// Cell Alignment submenu
			.addSubMenu(ui.createMenu("Cell Content Alignment")
				.addItem("Current Cell → Top", "menuAlignSelectedCellTop")
				.addItem("Current Cell → Middle", "menuAlignSelectedCellMiddle")
				.addItem("Current Cell → Bottom", "menuAlignSelectedCellBottom")
				.addSeparator()
				.addItem("All Cells → Top", "menuAlignAllCellsTop")
				.addItem("All Cells → Middle", "menuAlignAllCellsMiddle")
				.addItem("All Cells → Bottom", "menuAlignAllCellsBottom")
			)

			// Padding submenu
			.addSubMenu(ui.createMenu("Cell Padding")
				.addItem("Current Cell → 0pt", "menuSetPaddingCell0")
				.addItem("Current Cell → 5pt", "menuSetPaddingCell5")
				.addItem("Current Cell → 10pt", "menuSetPaddingCell10")
				.addItem("Current Cell → 20pt", "menuSetPaddingCell20")
				.addSeparator()
				.addItem("All Cells → 0pt", "menuSetPaddingTable0")
				.addItem("All Cells → 5pt", "menuSetPaddingTable5")
				.addItem("All Cells → 10pt", "menuSetPaddingTable10")
				.addItem("All Cells → 20pt", "menuSetPaddingTable20")
			)

			// Borders submenu
			.addSubMenu(ui.createMenu("Table Borders")
				.addItem("1pt Black Border", "menuSetTableBorder1ptBlack")
				.addItem("2pt Black Border", "menuSetTableBorder2ptBlack")
				.addItem("1pt Gray Border", "menuSetTableBorder1ptGray")
				.addItem("2pt Blue Border", "menuSetTableBorder2ptBlue")
				.addSeparator()
				.addItem("Remove All Borders", "menuRemoveTableBorders")
			)
		)

		.addSeparator()

		// HTML Operations
		.addSubMenu(ui.createMenu("🔧 HTML Export")
			.addItem("HTML to Start ↑", "menuDropHtmlStart")
			.addItem("HTML to End ↓", "menuDropHtmlEnd")
			.addItem("HTML to Clipboard", "menuDropHtmlClipboard")
			.addSeparator()
			.addItem("Strip HTML Tags", "menuStripHtmlTags")
			.addItem("Delete All HTML", "menuStripHtmlAll")
		)

		.addSeparator()

		// Background Colors
		.addSubMenu(ui.createMenu("🎨 Background")
			.addItem("Gray Background", "menuSetBackgroundGray")
			.addItem("White Background", "menuSetBackgroundWhite")
			.addItem("Dark Background", "menuSetBackgroundDark")
		)

		.addSeparator()

		// Testing & Development
		.addSubMenu(ui.createMenu("🧪 Testing")
			.addItem("Run Interactive Tests", "runInteractiveTests")
			.addItem("Test Table Operations", "testAllTableOperations")
			.addItem("Test Element Insertions", "testElementInsertions")
		)

		.addToUi();
}

function showFormInSidebar() {
	var form = HtmlService.createTemplateFromFile('Index').evaluate().setTitle('Pipewriter');
	var userProperties = PropertiesService.getUserProperties();
	DocumentApp.getUi().showSidebar(form);
}

// Legacy wrapper functions for backwards compatibility
function dropHtmlStart() {
	return dropHtml({ position: 'start' });
}

function dropHtmlEnd() {
	return dropHtml({ position: 'end' });
}

function dropHtmlClipboard() {
	return dropHtml({ copyToClipboard: true });
}

function stripHtmlTags() {
	return stripHtml({ all: false });
}

function stripHtmlAll() {
	return stripHtml({ all: true });
}

// Element getter wrapper - forwards to dropper.js
function getElement(elementId) {
	return dropper.getElement(elementId);
}

function insertZigzagRight() {
	return tableCreator.createZigzagRight();
}