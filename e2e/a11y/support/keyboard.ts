/**
 * The keyboard contract every surface has to keep, in a real browser where Tab order exists:
 *
 * - Tab from the page start comes back round to where it started, never sticking on one stop
 *   (three identical stops in a row is a trap), and passes through every region on the way;
 * - F6 and Shift+F6 from the document visit every region and come back;
 * - a dialog keeps Tab and Shift+Tab inside, closes on Escape and gives focus back;
 * - saving and applying an assistant edit leave focus where it was.
 */
import type { Locator, Page } from '@playwright/test';
import type { FocusStop } from './page-helpers';

/** A region of a product: a name and a selector that matches it or one of its ancestors. */
export type Region = readonly [name: string, selector: string];

export interface Contract {
	/** Regions, most specific first: a stop belongs to the first one it is inside. */
	regions: readonly Region[];
	/** The region that holds the document (the page, grid, slide, canvas or composer). */
	document: string;
	/** Focuses the document the way a user who has reached it would. */
	focusDocument(page: Page): Promise<void>;
	/** Regions the Tab cycle must pass through. */
	tabRegions: readonly string[];
	/** Regions F6 must reach from the document. */
	f6Regions: readonly string[];
	/**
	 * How the Tab walk leaves the document when Tab moves inside it, as in a spreadsheet grid
	 * where Tab goes to the next cell: the region key takes it on.
	 */
	leaveDocumentBy?: 'F6';
}

/** Where focus is now, in the contract's terms. */
export const focusStop = (page: Page, contract: Contract): Promise<FocusStop> =>
	page.evaluate(
		(regions) => globalThis.__a11y.stop(regions),
		contract.regions as [string, string][],
	);

const MAX_REPEATS = 2;

export interface TabCycle {
	closed: boolean;
	presses: number;
	regions: Set<string>;
	/** The stop focus stuck on, when it did. */
	trap?: string;
	/** A readable trail of the walk, for failure messages. */
	trail: string[];
}

/**
 * Presses Tab from the start of the page until focus is back on the first stop. A stop that holds
 * focus for three presses in a row is a trap and ends the walk.
 */
export async function tabCycle(page: Page, contract: Contract, max = 400): Promise<TabCycle> {
	await page.evaluate(() => {
		(globalThis.__a11y.deep() as HTMLElement | null)?.blur?.();
		getSelection()?.removeAllRanges();
	});
	const regions = new Set<string>();
	const trail: string[] = [];
	let first: FocusStop | undefined;
	let previous = -1;
	let repeats = 0;
	for (let presses = 1; presses <= max; presses += 1) {
		await page.keyboard.press('Tab');
		let stop = await focusStop(page, contract);
		if (stop.region === contract.document && contract.leaveDocumentBy) {
			regions.add(stop.region);
			trail.push(`${presses}: ${stop.label} <${stop.region}> then ${contract.leaveDocumentBy}`);
			await page.keyboard.press(contract.leaveDocumentBy);
			stop = await focusStop(page, contract);
		}
		trail.push(`${presses}: ${stop.label} <${stop.region ?? '-'}>`);
		if (stop.region) regions.add(stop.region);
		repeats = stop.id === previous ? repeats + 1 : 0;
		previous = stop.id;
		if (repeats >= MAX_REPEATS) return { closed: false, presses, regions, trap: stop.label, trail };
		if (!first) {
			if (stop.id !== 0) first = stop;
		} else if (stop.id === first.id) {
			return { closed: true, presses, regions, trail };
		}
	}
	return { closed: false, presses: max, regions, trail };
}

export interface RegionCycle {
	/** The regions in the order the key reached them, until it came back to the document. */
	order: string[];
	returned: boolean;
	trail: string[];
}

/** Presses `key` from the document until focus is back in it, recording each region reached. */
export async function regionCycle(
	page: Page,
	contract: Contract,
	key: 'F6' | 'Shift+F6',
): Promise<RegionCycle> {
	await contract.focusDocument(page);
	const order: string[] = [];
	const trail: string[] = [];
	const limit = contract.f6Regions.length * 2 + 2;
	for (let press = 0; press < limit; press += 1) {
		await page.keyboard.press(key);
		await page.waitForTimeout(50);
		const stop = await focusStop(page, contract);
		trail.push(`${stop.label} <${stop.region ?? '-'}>`);
		if (stop.region === contract.document) return { order, returned: true, trail };
		order.push(stop.region ?? stop.label);
	}
	return { order, returned: false, trail };
}

const OPEN_DIALOG = 'office-ui-dialog[open], dialog[open], [role="dialog"], [role="alertdialog"]';

/** Whether focus is inside an open dialog. */
export const focusInDialog = (page: Page): Promise<boolean> =>
	page.evaluate((selector) => globalThis.__a11y.focusWithin(selector), OPEN_DIALOG);

/** The dialogs showing now. */
const openDialogs = (page: Page): Promise<number[]> =>
	page.evaluate((selector) => globalThis.__a11y.visible(selector), OPEN_DIALOG);

/** How a keyboard user opens one of the product's dialogs, and where focus belongs after it. */
export interface DialogOpener {
	/** Gets the page ready, for example by showing a ribbon tab or focusing the document. */
	prepare?(page: Page): Promise<void>;
	/** The control that opens it; without one, `key` is pressed wherever `prepare` left focus. */
	opener?(page: Page): Locator;
	/** The key that opens it (Enter on the opener by default). */
	key?: string;
	/** Where focus belongs once it closes: back on its opener, or in the document. */
	returnsTo: 'opener' | 'document';
}

export interface DialogRun {
	focusMovedIn: boolean;
	/** Presses after which focus was outside the dialog, as `Tab 3` or `Shift+Tab 1`. */
	escapes: string[];
	closedByEscape: boolean;
	/** Where focus was before the dialog opened, and where it went after Escape. */
	before: FocusStop;
	after: FocusStop;
}

/** Opens a dialog from the keyboard, walks Tab and Shift+Tab inside it, then presses Escape. */
export async function dialogRun(
	page: Page,
	contract: Contract,
	dialog: DialogOpener,
	presses = 25,
): Promise<DialogRun> {
	await dialog.prepare?.(page);
	await dialog.opener?.(page).focus();
	const before = await focusStop(page, contract);
	const showing = new Set(await openDialogs(page));
	await page.keyboard.press(dialog.key ?? 'Enter');
	await page
		.waitForFunction((selector) => globalThis.__a11y.focusWithin(selector), OPEN_DIALOG, {
			timeout: 5_000,
		})
		.catch(() => undefined);
	const focusMovedIn = await focusInDialog(page);
	const escapes: string[] = [];
	// A native modal dialog lets Tab pass through the browser's own controls, where no page
	// element has focus; the next press shows whether focus comes back in or lands behind it.
	const where = () =>
		page.evaluate(
			(selector) =>
				globalThis.__a11y.focusWithin(selector) ? 'in' : globalThis.__a11y.deep() ? 'out' : 'none',
			OPEN_DIALOG,
		);
	for (const key of ['Tab', 'Shift+Tab'] as const)
		for (let press = 1; press <= presses; press += 1) {
			await page.keyboard.press(key);
			if ((await where()) === 'out') escapes.push(`${key} ${press}`);
		}
	await page.keyboard.press('Escape');
	await page.waitForTimeout(300);
	return {
		focusMovedIn,
		escapes,
		// Only the dialog this run opened: a panel that is always a dialog may stay.
		closedByEscape: (await openDialogs(page)).every((id) => showing.has(id)),
		before,
		after: await focusStop(page, contract),
	};
}
