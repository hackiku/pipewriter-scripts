// utils/menuWrappers.gs

// --- Wrappers for Element Insertion (dropper.js) ---
function menuInsertZigzagLeft() {
	// Assumes 'dropper' is global from dropper.js
	// showAlert is implicitly true as dropper.getElement uses showUserError
	var result = dropper.getElement({ elementId: 'zz-left', theme: 'light' });
	// dropper.getElement already shows an alert on error, so no need to repeat here
	// if (!result.success) DocumentApp.getUi().alert('Error: ' + result.error); 
}

function menuInsertBlurbs3() {
	var result = dropper.getElement({ elementId: 'blurbs-3', theme: 'light' });
	// if (!result.success) DocumentApp.getUi().alert('Error: ' + result.error);
}

// --- Wrapper for Programmatic Table Creation (createTable.js) ---
function menuInsertZigzagRight() {
	// Assumes 'tableCreator' is global from createTable.js
	var result = tableCreator.createZigzagRight();
	if (!result.success) DocumentApp.getUi().alert('Error creating Zigzag Right: ' + result.error);
}

// --- Wrappers for HTML Operations (convertHtml.js) ---
function menuDropHtmlStart() {
	// Assumes 'dropHtml' is global from convertHtml.js
	var result = dropHtml({ position: 'start' });
	if (!result.success) DocumentApp.getUi().alert('Error formatting to HTML (start): ' + result.error);
}

function menuDropHtmlEnd() {
	var result = dropHtml({ position: 'end' });
	if (!result.success) DocumentApp.getUi().alert('Error formatting to HTML (end): ' + result.error);
}

function menuDropHtmlClipboard() {
	var result = dropHtml({ copyToClipboard: true });
	// copyHtmlToClipboard in convertHtml.js shows a modal or an alert on failure.
	// No need for additional alert here if dropHtml handles it or returns specific status.
}

function menuStripHtmlTags() {
	// Assumes 'stripHtml' is global
	var result = stripHtml({ all: false, copyToClipboard: false }); // Default is tags only
	if (result.success) {
		const message = result.message || (result.replacedCount !== undefined ? `Removed ${result.replacedCount} HTML tags.` : 'HTML tags stripped.');
		DocumentApp.getUi().alert(message);
	} else {
		DocumentApp.getUi().alert('Error stripping HTML tags: ' + result.error);
	}
}

function menuStripHtmlAll() {
	var result = stripHtml({ all: true, copyToClipboard: false });
	if (result.success) {
		const message = result.message || (result.removedCount !== undefined ? `Removed ${result.removedCount} HTML elements.` : 'HTML elements deleted.');
		DocumentApp.getUi().alert(message);
	} else {
		DocumentApp.getUi().alert('Error deleting HTML: ' + result.error);
	}
}

// --- Wrapper for Color Operations (color.js) ---
function menuGrayBackground() {
	// Assumes 'changeBg' is global from color.js
	var result = changeBg('#f3f3f3');
	if (!result.success) DocumentApp.getUi().alert('Error changing background: ' + result.error);
}


// --- Wrappers for Table Operations (formatting/table.js) ---
// These call the global tableOps(params) which is defined in formatting/table.js
// The showAlert default in tableOps is true, so UI alerts will show for these menu actions.

// Cell Content Alignment - Selected Cell
function menuAlignSelectedCellTop() { tableOps({ action: 'setCellAlignment', scope: 'cell', alignment: 'top' }); }
function menuAlignSelectedCellMiddle() { tableOps({ action: 'setCellAlignment', scope: 'cell', alignment: 'middle' }); }
function menuAlignSelectedCellBottom() { tableOps({ action: 'setCellAlignment', scope: 'cell', alignment: 'bottom' }); }

// Cell Content Alignment - Whole Table
function menuAlignAllCellsTop() { tableOps({ action: 'setCellAlignment', scope: 'table', alignment: 'top' }); }
function menuAlignAllCellsMiddle() { tableOps({ action: 'setCellAlignment', scope: 'table', alignment: 'middle' }); }
function menuAlignAllCellsBottom() { tableOps({ action: 'setCellAlignment', scope: 'table', alignment: 'bottom' }); }

// Padding - Selected Cell
function menuSetPaddingCell0() { tableOps({ action: 'setCellPadding', scope: 'cell', padding: 0 }); }
function menuSetPaddingCell5() { tableOps({ action: 'setCellPadding', scope: 'cell', padding: 5 }); }
function menuSetPaddingCell10() { tableOps({ action: 'setCellPadding', scope: 'cell', padding: 10 }); }
function menuSetPaddingCell20() { tableOps({ action: 'setCellPadding', scope: 'cell', padding: 20 }); }

// Padding - Whole Table
function menuSetPaddingTable0() { tableOps({ action: 'setCellPadding', scope: 'table', padding: 0 }); }
function menuSetPaddingTable5() { tableOps({ action: 'setCellPadding', scope: 'table', padding: 5 }); }
function menuSetPaddingTable10() { tableOps({ action: 'setCellPadding', scope: 'table', padding: 10 }); }
function menuSetPaddingTable20() { tableOps({ action: 'setCellPadding', scope: 'table', padding: 20 }); }

// Borders - Whole Table
function menuSetTableBorder1ptBlack() { tableOps({ action: 'setBorders', scope: 'table', borderWidth: 1, borderColor: '#000000' }); }
function menuSetTableBorder2ptBlue() { tableOps({ action: 'setBorders', scope: 'table', borderWidth: 2, borderColor: '#0000FF' }); }
function menuRemoveTableBorders() { tableOps({ action: 'setBorders', scope: 'table', borderWidth: 0 }); }

// Select Table
function menuSelectTable() { tableOps({ action: 'selectWholeTable' }); }