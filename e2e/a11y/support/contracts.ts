/**
 * Each product's regions and how a keyboard user reaches its document and a dialog. The suite hosts
 * Word, Excel, PowerPoint (the vanilla binding), Visio and Teams; the PowerPoint demos host the
 * other bindings with the same contract.
 */
import type { Page } from '@playwright/test';
import type { Contract, DialogOpener, Region } from './keyboard';

export interface ProductContract extends Contract {
	title: string;
	dialog: DialogOpener;
	/** One edit made from the keyboard, for the save check. */
	edit?(page: Page): Promise<void>;
}

/** The suite's own chrome around every editor. */
const WORKSPACE: Region = ['workspace', '.suite-header, .workspace-sidebar, .tab-row'];

const focus = (selector: string) => async (page: Page) => {
	await page.locator(selector).first().focus();
};
const focusGrid = (page: Page) =>
	page
		.locator('xlsx-editor')
		.evaluate((editor) => (editor as HTMLElement & { focusGrid(): void }).focusGrid());

/** PowerPoint's panes, marked the same way in every binding (the panes F6 visits). */
export const POWERPOINT_REGIONS: readonly Region[] = [
	['slide', '[data-pptx-viewport]'],
	['thumbnails', '[data-pptx-chrome="slides"]'],
	['notes', '[data-pptx-chrome="notes"]'],
	['format pane', '[data-pptx-chrome="inspector"]'],
	['status', 'pptx-ui-status-bar, [data-pptx-chrome="status-bar"]'],
	['ribbon', '[data-pptx-chrome="ribbon"]'],
];

export const powerPoint = (regions: readonly Region[], tabRegions: readonly string[]) =>
	({
		title: 'PowerPoint',
		regions,
		document: 'slide',
		focusDocument: focus('[data-pptx-viewport]'),
		tabRegions,
		f6Regions: ['ribbon', 'thumbnails', 'notes', 'format pane', 'status'],
		// The Settings dialog, which the vanilla binding calls Settings & Shortcuts.
		dialog: {
			opener: (page) => page.getByRole('button', { name: /^Settings( & Shortcuts)?$/u }).first(),
			returnsTo: 'opener',
		},
		async edit(page) {
			// Enter goes into the slide and selects the first object; an arrow key nudges it.
			await page.locator('[data-pptx-viewport]').first().focus();
			await page.keyboard.press('Enter');
			await page.keyboard.press('ArrowRight');
		},
	}) satisfies ProductContract;

export type SuiteProduct = 'docx' | 'xlsx' | 'pptx' | 'vsdx' | 'teams';

export const SUITE: Record<SuiteProduct, ProductContract> = {
	docx: {
		title: 'Word',
		regions: [
			['document', '.ProseMirror'],
			['ribbon', 'office-ui-ribbon'],
			['status', 'office-ui-status-bar'],
			WORKSPACE,
		],
		document: 'document',
		focusDocument: focus('docx-editor .ProseMirror'),
		tabRegions: ['workspace', 'ribbon', 'document', 'status'],
		f6Regions: ['ribbon', 'status'],
		// As in Word, a dialog opened from the ribbon hands the keyboard back to the document.
		dialog: {
			opener: (page) => page.getByRole('button', { name: 'Font settings', exact: true }),
			returnsTo: 'document',
		},
		async edit(page) {
			await page.locator('docx-editor .ProseMirror').first().focus();
			await page.keyboard.press('Control+End');
			await page.keyboard.type('Q');
		},
	},
	xlsx: {
		title: 'Excel',
		regions: [
			// The focus sink lives in the grid's root, beside the accessible grid of cells.
			['grid', '[part="grid"]'],
			['formula bar', '[part="formula-bar"]'],
			['sheet tabs', '[part="sheet-tabs"]'],
			['ribbon', 'office-ui-ribbon'],
			['status', 'office-ui-status-bar'],
			WORKSPACE,
		],
		document: 'grid',
		focusDocument: focusGrid,
		tabRegions: ['workspace', 'ribbon', 'formula bar', 'grid', 'sheet tabs', 'status'],
		f6Regions: ['ribbon', 'formula bar', 'sheet tabs', 'status'],
		// Tab moves to the next cell, as in Excel; F6 is the way on.
		leaveDocumentBy: 'F6',
		dialog: { prepare: focusGrid, key: 'Control+1', returnsTo: 'document' },
		async edit(page) {
			await focusGrid(page);
			await page.keyboard.press('Control+End');
			await page.keyboard.press('ArrowDown');
			await page.keyboard.type('5');
			await page.keyboard.press('Enter');
		},
	},
	pptx: powerPoint(
		[...POWERPOINT_REGIONS, WORKSPACE],
		['workspace', 'ribbon', 'thumbnails', 'slide', 'format pane', 'notes', 'status'],
	),
	vsdx: {
		title: 'Visio',
		regions: [
			['canvas', '[aria-label="Diagram canvas"]'],
			['page tabs', '[role="tablist"][aria-label="Pages"]'],
			['shapes', '#shapes-pane'],
			['ribbon', 'office-ui-ribbon'],
			WORKSPACE,
		],
		document: 'canvas',
		focusDocument: focus('[aria-label="Diagram canvas"]'),
		tabRegions: ['workspace', 'ribbon', 'shapes', 'canvas', 'page tabs'],
		f6Regions: ['ribbon', 'shapes', 'page tabs'],
		dialog: {
			prepare: (page) => page.getByRole('tab', { name: 'Review', exact: true }).last().click(),
			opener: (page) => page.getByRole('button', { name: 'Page Alt Text', exact: true }),
			returnsTo: 'opener',
		},
		async edit(page) {
			// Tab from the canvas focuses the first shape, Space selects it, an arrow key nudges it.
			await page.locator('[aria-label="Diagram canvas"]').first().focus();
			await page.keyboard.press('Tab');
			await page.keyboard.press('Space');
			await page.keyboard.press('ArrowRight');
		},
	},
	teams: {
		title: 'Teams',
		regions: [
			['composer', 'textarea[aria-label="Message"]'],
			['app bar', '[aria-label="App bar"]'],
			['channels', 'office-ui-channel-list'],
			['channel header', 'header.channel-head'],
			WORKSPACE,
		],
		document: 'composer',
		focusDocument: focus('teams-app textarea[aria-label="Message"]'),
		tabRegions: ['workspace', 'app bar', 'channels', 'channel header', 'composer'],
		f6Regions: ['app bar', 'channels', 'channel header'],
		dialog: {
			opener: (page) => page.getByRole('button', { name: 'Add channel', exact: true }),
			returnsTo: 'opener',
		},
	},
};
