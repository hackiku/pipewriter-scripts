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
		DocumentApp.getUi().alert('❌ Error', 'Please provide a valid padding value (0 or greater)', DocumentApp.getUi().ButtonSet.OK);
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
		DocumentApp.getUi().alert('❌ Error', 'Please provide a valid padding value (0 or greater)', DocumentApp.getUi().ButtonSet.OK);
		return { success: false, error: 'Invalid padding value' };
	}
	return tableOps({ action: 'setCellPadding', scope: 'table', padding: padding });
}

// --- Table Borders ---
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

// Custom border for table
function menuSetTableBorderCustom(borderWidth, borderColor = '#000000') {
	if (typeof borderWidth !== 'number' || borderWidth < 0) {
		DocumentApp.getUi().alert('❌ Error', 'Please provide a valid border width (0 or greater)', DocumentApp.getUi().ButtonSet.OK);
		return { success: false, error: 'Invalid border width' };
	}
	return tableOps({ action: 'setBorders', scope: 'table', borderWidth: borderWidth, borderColor: borderColor });
}

// --- Table Selection & Options ---
function menuSelectTable() {
	return tableOps({ action: 'selectWholeTable' });
}

function menuOpenTableOptions() {
	return tableOps({ action: 'openTableOptions' });
}

// =============================================================================
// TESTING UTILITIES
// =============================================================================

/**
 * Test all table operations to verify they work
 * Place cursor in a table cell before running
 */
function testAllTableOperations() {
	Logger.log('=== Testing Table Operations ===');

	const tests = [
		{ name: 'Align Cell Top', fn: () => menuAlignSelectedCellTop() },
		{ name: 'Align Cell Middle', fn: () => menuAlignSelectedCellMiddle() },
		{ name: 'Set Cell Padding 10pt', fn: () => menuSetPaddingCell10() },
		{ name: 'Set Table Padding 5pt', fn: () => menuSetPaddingTable5() },
		{ name: 'Add 1pt Black Border', fn: () => menuSetTableBorder1ptBlack() },
		{ name: 'Remove Borders', fn: () => menuRemoveTableBorders() },
		{ name: 'Select Table', fn: () => menuSelectTable() }
	];

	const results = [];

	tests.forEach(test => {
		try {
			Logger.log(`Testing: ${test.name}`);
			const result = test.fn();
			results.push({
				name: test.name,
				success: result.success,
				message: result.message || result.error,
				executionTime: result.executionTime
			});
			Logger.log(`✅ ${test.name}: ${result.success ? 'PASSED' : 'FAILED'}`);
			if (!result.success) {
				Logger.log(`   Error: ${result.error}`);
			}
		} catch (error) {
			Logger.log(`❌ ${test.name}: ERROR - ${error.toString()}`);
			results.push({
				name: test.name,
				success: false,
				message: error.toString(),
				executionTime: null
			});
		}
	});

	Logger.log('=== Test Results Summary ===');
	const passed = results.filter(r => r.success).length;
	const total = results.length;
	Logger.log(`Passed: ${passed}/${total}`);

	return results;
}

/**
 * Test element insertion operations
 */
function testElementInsertions() {
	Logger.log('=== Testing Element Insertions ===');

	const tests = [
		{ name: 'Insert Hero', fn: () => menuInsertHero() },
		{ name: 'Insert Zigzag Left', fn: () => menuInsertZigzagLeft() },
		{ name: 'Insert Blurbs 3', fn: () => menuInsertBlurbs3() }
	];

	const results = [];

	tests.forEach(test => {
		try {
			Logger.log(`Testing: ${test.name}`);
			const result = test.fn();
			results.push({
				name: test.name,
				success: result.success,
				message: result.message || result.error
			});
			Logger.log(`${result.success ? '✅' : '❌'} ${test.name}: ${result.success ? 'PASSED' : 'FAILED'}`);
		} catch (error) {
			Logger.log(`❌ ${test.name}: ERROR - ${error.toString()}`);
			results.push({
				name: test.name,
				success: false,
				message: error.toString()
			});
		}
	});

	return results;
}

/**
 * Interactive test runner - shows results in UI
 */
function runInteractiveTests() {
	const ui = DocumentApp.getUi();

	const response = ui.alert(
		'Pipewriter Test Runner',
		'What would you like to test?\n\n1. Table Operations (requires cursor in table cell)\n2. Element Insertions (requires cursor in document)',
		ui.ButtonSet.YES_NO_CANCEL
	);

	let results = [];

	if (response === ui.Button.YES) {
		results = testAllTableOperations();
	} else if (response === ui.Button.NO) {
		results = testElementInsertions();
	} else {
		return;
	}

	// Show results
	const passed = results.filter(r => r.success).length;
	const total = results.length;
	const summary = `Test Results: ${passed}/${total} passed\n\n` +
		results.map(r => `${r.success ? '✅' : '❌'} ${r.name}`).join('\n');

	ui.alert('Test Results', summary, ui.ButtonSet.OK);
}