// utils/menuWrappers.js - Clean, testable menu wrapper functions

// =============================================================================
// ELEMENT INSERTION WRAPPERS (dropper.js)
// =============================================================================

function menuInsertZigzagLeft() {
	return dropper.getElement({ elementId: 'zz-left', theme: 'light' });
}

function menuInsertZigzagRight() {
	return tableCreator.createZigzagRight();
}

function menuInsertBlurbs3() {
	return dropper.getElement({ elementId: 'blurbs-3', theme: 'light' });
}

function menuInsertHero() {
	return dropper.getElement({ elementId: 'hero', theme: 'light' });
}

function menuInsertCards2() {
	return dropper.getElement({ elementId: 'cards-2', theme: 'light' });
}

function menuInsertCards3() {
	return dropper.getElement({ elementId: 'cards-3', theme: 'light' });
}

// =============================================================================
// TEXT OPERATIONS WRAPPERS (formatting/text.js)
// =============================================================================

function menuApplyNormal() {
	return textOps({ action: 'applyStyle', headingType: 'NORMAL' });
}

function menuApplyHeading1() {
	return textOps({ action: 'applyStyle', headingType: 'HEADING1' });
}

function menuApplyHeading2() {
	return textOps({ action: 'applyStyle', headingType: 'HEADING2' });
}

function menuApplyHeading3() {
	return textOps({ action: 'applyStyle', headingType: 'HEADING3' });
}

function menuApplyHeading4() {
	return textOps({ action: 'applyStyle', headingType: 'HEADING4' });
}

function menuApplyHeading5() {
	return textOps({ action: 'applyStyle', headingType: 'HEADING5' });
}

function menuApplyHeading6() {
	return textOps({ action: 'applyStyle', headingType: 'HEADING6' });
}

function menuUpdateAllMatching() {
	return textOps({ action: 'updateAllMatching' });
}

// Menu version shows alert popup (for dropdown submenu)
function menuGetStyleInfo() {
	const startTime = new Date().getTime();

	try {
		const context = getTextContext();
		if (!context.success) {
			throw new Error(context.error);
		}

		// Call the version WITH alert
		return getStyleInfoWithAlert(context, startTime);
	} catch (error) {
		const errorResult = {
			success: false,
			error: error.toString(),
			message: error.message || 'Please place cursor in document.',
			executionTime: new Date().getTime() - startTime
		};

		Logger.log(`Error in menuGetStyleInfo: ${error.message}`);
		return errorResult;
	}
}

// =============================================================================
// HTML OPERATIONS WRAPPERS (ai/convertHtml.js)
// =============================================================================

function menuDropHtmlStart() {
	return dropHtml({ position: 'start' });
}

function menuDropHtmlEnd() {
	return dropHtml({ position: 'end' });
}

function menuDropHtmlClipboard() {
	return dropHtml({ copyToClipboard: true });
}

function menuStripHtmlTags() {
	return stripHtml({ all: false });
}

function menuStripHtmlAll() {
	return stripHtml({ all: true });
}

// =============================================================================
// COLOR OPERATIONS WRAPPERS (formatting/color.js)
// =============================================================================

function menuSetBackgroundGray() {
	return changeBg('#f3f3f3');
}

function menuSetBackgroundWhite() {
	return changeBg('#ffffff');
}

function menuSetBackgroundDark() {
	return changeBg('#2c2c2c');
}

// =============================================================================
// TABLE OPERATIONS WRAPPERS (formatting/table.js)
// =============================================================================

// --- Cell Content Alignment (Selected Cell) ---
function menuAlignSelectedCellTop() {
	return tableOps({ action: 'setCellAlignment', scope: 'cell', alignment: 'top' });
}

function menuAlignSelectedCellMiddle() {
	return tableOps({ action: 'setCellAlignment', scope: 'cell', alignment: 'middle' });
}

function menuAlignSelectedCellBottom() {
	return tableOps({ action: 'setCellAlignment', scope: 'cell', alignment: 'bottom' });
}

// --- Cell Content Alignment (Whole Table) ---
function menuAlignAllCellsTop() {
	return tableOps({ action: 'setCellAlignment', scope: 'table', alignment: 'top' });
}

function menuAlignAllCellsMiddle() {
	return tableOps({ action: 'setCellAlignment', scope: 'table', alignment: 'middle' });
}

function menuAlignAllCellsBottom() {
	return tableOps({ action: 'setCellAlignment', scope: 'table', alignment: 'bottom' });
}

// --- Padding (Selected Cell) ---
function menuSetPaddingCell0() {
	return tableOps({ action: 'setCellPadding', scope: 'cell', padding: 0 });
}

function menuSetPaddingCell5() {
	return tableOps({ action: 'setCellPadding', scope: 'cell', padding: 5 });
}

function menuSetPaddingCell10() {
	return tableOps({ action: 'setCellPadding', scope: 'cell', padding: 10 });
}

function menuSetPaddingCell20() {
	return tableOps({ action: 'setCellPadding', scope: 'cell', padding: 20 });
}

// Custom padding for selected cell
function menuSetPaddingCellCustom(padding) {
	if (typeof padding !== 'number' || padding < 0) {
		return { success: false, error: 'Invalid padding value' };
	}
	return tableOps({ action: 'setCellPadding', scope: 'cell', padding: padding });
}

// --- Padding (Whole Table) ---
function menuSetPaddingTable0() {
	return tableOps({ action: 'setCellPadding', scope: 'table', padding: 0 });
}

function menuSetPaddingTable5() {
	return tableOps({ action: 'setCellPadding', scope: 'table', padding: 5 });
}

function menuSetPaddingTable10() {
	return tableOps({ action: 'setCellPadding', scope: 'table', padding: 10 });
}

function menuSetPaddingTable20() {
	return tableOps({ action: 'setCellPadding', scope: 'table', padding: 20 });
}

// Custom padding for whole table
function menuSetPaddingTableCustom(padding) {
	if (typeof padding !== 'number' || padding < 0) {
		return { success: false, error: 'Invalid padding value' };
	}
	return tableOps({ action: 'setCellPadding', scope: 'table', padding: padding });
}

// --- Table Borders (Table-wide only) ---
function menuSetTableBorder1ptBlack() {
	return tableOps({ action: 'setBorders', scope: 'table', borderWidth: 1, borderColor: '#000000' });
}

function menuSetTableBorder2ptBlack() {
	return tableOps({ action: 'setBorders', scope: 'table', borderWidth: 2, borderColor: '#000000' });
}

function menuSetTableBorder1ptGray() {
	return tableOps({ action: 'setBorders', scope: 'table', borderWidth: 1, borderColor: '#cccccc' });
}

function menuSetTableBorder2ptBlue() {
	return tableOps({ action: 'setBorders', scope: 'table', borderWidth: 2, borderColor: '#0000FF' });
}

function menuRemoveTableBorders() {
	return tableOps({ action: 'setBorders', scope: 'table', borderWidth: 0 });
}

// Custom table border
function menuSetTableBorderCustom(borderWidth, borderColor = '#000000') {
	if (typeof borderWidth !== 'number' || borderWidth < 0) {
		return { success: false, error: 'Invalid border width' };
	}
	return tableOps({ action: 'setBorders', scope: 'table', borderWidth: borderWidth, borderColor: borderColor });
}

// --- Cell Background Colors (Current Cell) ---
function menuSetCellBackgroundWhite() {
	return tableOps({ action: 'setCellBackground', scope: 'cell', backgroundColor: '#ffffff' });
}

function menuSetCellBackgroundLightGray() {
	return tableOps({ action: 'setCellBackground', scope: 'cell', backgroundColor: '#f3f3f3' });
}

function menuSetCellBackgroundDarkGray() {
	return tableOps({ action: 'setCellBackground', scope: 'cell', backgroundColor: '#e0e0e0' });
}

function menuSetCellBackgroundBlue() {
	return tableOps({ action: 'setCellBackground', scope: 'cell', backgroundColor: '#e3f2fd' });
}

function menuSetCellBackgroundGreen() {
	return tableOps({ action: 'setCellBackground', scope: 'cell', backgroundColor: '#e8f5e8' });
}

function menuSetCellBackgroundYellow() {
	return tableOps({ action: 'setCellBackground', scope: 'cell', backgroundColor: '#fff9c4' });
}

function menuClearCellBackground() {
	return tableOps({ action: 'setCellBackground', scope: 'cell', backgroundColor: '' });
}

// --- Table Background Colors (All Cells) ---
function menuSetTableBackgroundWhite() {
	return tableOps({ action: 'setCellBackground', scope: 'table', backgroundColor: '#ffffff' });
}

function menuSetTableBackgroundLightGray() {
	return tableOps({ action: 'setCellBackground', scope: 'table', backgroundColor: '#f3f3f3' });
}

function menuSetTableBackgroundDarkGray() {
	return tableOps({ action: 'setCellBackground', scope: 'table', backgroundColor: '#e0e0e0' });
}

function menuSetTableBackgroundBlue() {
	return tableOps({ action: 'setCellBackground', scope: 'table', backgroundColor: '#e3f2fd' });
}

function menuSetTableBackgroundGreen() {
	return tableOps({ action: 'setCellBackground', scope: 'table', backgroundColor: '#e8f5e8' });
}

function menuSetTableBackgroundYellow() {
	return tableOps({ action: 'setCellBackground', scope: 'table', backgroundColor: '#fff9c4' });
}

function menuClearTableBackground() {
	return tableOps({ action: 'setCellBackground', scope: 'table', backgroundColor: '' });
}

// Custom background colors
function menuSetCellBackgroundCustom(backgroundColor) {
	if (!backgroundColor) {
		return { success: false, error: 'No background color specified' };
	}
	return tableOps({ action: 'setCellBackground', scope: 'cell', backgroundColor: backgroundColor });
}

function menuSetTableBackgroundCustom(backgroundColor) {
	if (!backgroundColor) {
		return { success: false, error: 'No background color specified' };
	}
	return tableOps({ action: 'setCellBackground', scope: 'table', backgroundColor: backgroundColor });
}

// --- Table Selection ---
function menuSelectTable() {
	return tableOps({ action: 'selectWholeTable' });
}