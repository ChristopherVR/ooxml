/**
 * The README recordings: one scene per picture. `setup` gets the page ready off camera (the
 * recording is cut where it returns); `play` is what the GIF shows. `out` is the GIF's path from
 * the repository root, beside the still picture it replaces.
 */

/** Light theme and the ribbon layout, like the product pictures in the docs sites. */
function productInit() {
	globalThis.ooxmlSkipStart = true;
	try {
		localStorage.setItem('ooxml-suite-theme', JSON.stringify({ mode: 'light', accent: 'blue' }));
		localStorage.setItem('ooxml-suite-settings', JSON.stringify({ layout: 'office' }));
	} catch {
		// Storage is unavailable: the app falls back to its defaults.
	}
}

/** The suite as a first visit sees it (dark), past the start screen. */
function suiteInit() {
	globalThis.ooxmlSkipStart = true;
}

let pointer = { x: 640, y: 400 };
/** Glide the mouse to a point, so the recording shows where it goes. */
async function glide(page, x, y, ms = 450) {
	const steps = Math.max(4, Math.round(ms / 16));
	await page.mouse.move(x, y, { steps });
	pointer = { x, y };
}
async function center(locator) {
	const box = await locator.boundingBox();
	if (!box) throw new Error(`Not visible: ${locator}`);
	return { x: box.x + box.width / 2, y: box.y + box.height / 2 };
}
/** Move to the middle of an element and click it. */
async function clickOn(page, locator, { pause = 150, ms } = {}) {
	await locator.waitFor({ state: 'visible' });
	const { x, y } = await center(locator);
	await glide(page, x, y, ms);
	await page.waitForTimeout(pause);
	await page.mouse.down();
	await page.mouse.up();
}
async function ready(page) {
	await page.waitForFunction(() => document.querySelector('#status')?.textContent === 'Ready');
}
/** Open a sample from the gallery's link and wait until its editor has drawn. */
async function openSample(page, base, app, id) {
	await page.goto(`${base}${app}?sample=${id}`);
	await page.locator('#tabs [data-open]').filter({ hasText: id }).first().waitFor();
	await ready(page);
}
const type = (page, text, delay = 55) => page.keyboard.type(text, { delay });

export const SCENES = [
	{
		name: 'chooser',
		out: '.github/assets/suite-chooser.gif',
		width: 800,
		viewport: { width: 1280, height: 800 },
		async setup(page, base) {
			await page.goto(base);
			await page.locator('#start-screen [data-start="office"]').waitFor();
			await page.waitForTimeout(1200);
			pointer = { x: 640, y: 120 };
			await page.mouse.move(pointer.x, pointer.y);
		},
		async play(page) {
			for (const id of ['docx', 'xlsx', 'pptx', 'vsdx', 'teams']) {
				const card = page.locator(`#start-screen [data-start="${id}"]`);
				const { x, y } = await center(card);
				await glide(page, x, y, 420);
				await page.waitForTimeout(380);
			}
			await page.waitForTimeout(200);
			await clickOn(page, page.locator('#start-screen [data-start="office"]'), { ms: 500 });
			// The suite opens on Chat; the file home is one click away.
			await page.locator('#chat-input').waitFor();
			await page.waitForTimeout(900);
			await clickOn(page, page.locator('#rail [data-home]').first());
			await page.locator('[data-create="docx"]').first().waitFor();
			await page.waitForTimeout(1500);
		},
	},
	{
		name: 'workspace',
		out: '.github/assets/suite-workspace.gif',
		init: suiteInit,
		width: 760,
		fps: 8,
		async setup(page, base) {
			await openSample(page, base, '/', 'sales-analysis');
			await openSample(page, base, '/', 'product-launch-keynote');
			await openSample(page, base, '/', 'order-to-delivery-flowchart');
			await page.goto(`${base}/#/`);
			await page.locator('[data-create="docx"]').first().waitFor();
			await page.waitForTimeout(1500);
			await page.mouse.move(700, 300);
		},
		async play(page) {
			const home = page.locator('#tabs [data-home]').first();
			const tab = (name) => page.locator('#tabs [data-open]').filter({ hasText: name }).first();
			/** Scroll Home with the wheel until the target is on screen, then click it. */
			const reveal = async (target, options) => {
				await target.waitFor();
				await glide(page, 760, 520, 300);
				for (let i = 0; i < 12; i++) {
					const box = await target.boundingBox();
					if (box && box.y > 100 && box.y + box.height < 760) break;
					await page.mouse.wheel(0, box && box.y <= 100 ? -120 : 120);
					await page.waitForTimeout(90);
				}
				await clickOn(page, target, options);
			};
			const openFile = (name, options) =>
				reveal(page.locator('#files [data-open]').filter({ hasText: name }).first(), options);
			await page.waitForTimeout(500);
			// Open two files from the library; each gets its own live tab.
			await openFile('sales-analysis', { ms: 500 });
			await page.locator('.document-host:not([hidden]) xlsx-editor').waitFor();
			await page.waitForTimeout(1500);
			await clickOn(page, home);
			await page.waitForTimeout(500);
			await openFile('product-launch-keynote');
			await page.waitForTimeout(1800);
			// The open documents stay live in their tabs.
			await clickOn(page, tab('sales-analysis'));
			await page.waitForTimeout(1100);
			await clickOn(page, tab('order-to-delivery-flowchart'));
			await page.waitForTimeout(1300);
		},
	},
	{
		name: 'word',
		out: 'viewers/docx/docs/public/hero-editor.gif',
		init: productInit,
		async setup(page, base) {
			await page.goto(`${base}/apps/word/`);
			await page.locator('[data-create="docx"]').first().click();
			await page.locator('docx-editor .ProseMirror[contenteditable=true]').first().waitFor();
			await ready(page);
			// Print Layout, as Word opens a document: pages on a canvas, thumbnails in the Pages pane.
			const print = page
				.locator('docx-editor')
				.getByRole('button', { name: 'Print Layout' })
				.first();
			if (await print.isVisible().catch(() => false)) await print.click();
			await page.waitForTimeout(1500);
			await page.mouse.move(700, 600);
		},
		async play(page) {
			const word = page.locator('docx-editor');
			const button = (name) => word.getByRole('button', { name, exact: true }).first();
			const editor = page.locator('docx-editor .ProseMirror[contenteditable=true]').first();
			await clickOn(page, editor);
			await type(page, 'Spring planting plan', 50);
			await page.keyboard.press('Enter');
			await type(
				page,
				'Sow tomatoes and peppers indoors in early March, then plant them out in May.',
				28,
			);
			await page.waitForTimeout(300);
			// Format the title: select it, then Bold and a larger size from the ribbon.
			await page.keyboard.press('Control+Home');
			await page.keyboard.press('Shift+End');
			await page.waitForTimeout(300);
			await clickOn(page, button('Bold'));
			await page.waitForTimeout(250);
			for (let i = 0; i < 3; i++) {
				await clickOn(page, button('Grow font'), { ms: i ? 120 : 450, pause: 100 });
				await page.waitForTimeout(200);
			}
			// Highlight two words of the body.
			await page.keyboard.press('Control+End');
			for (let i = 0; i < 7; i++) await page.keyboard.press('Control+ArrowLeft');
			await page.keyboard.press('Control+Shift+ArrowRight');
			await page.keyboard.press('Control+Shift+ArrowRight');
			await page.waitForTimeout(300);
			await clickOn(page, button('Text highlight'));
			await page.waitForTimeout(300);
			await clickOn(page, button('Italic'));
			await page.waitForTimeout(1600);
		},
	},
	{
		name: 'excel',
		out: 'viewers/xlsx/docs/public/hero-editor.gif',
		init: productInit,
		async setup(page, base) {
			await page.goto(`${base}/apps/excel/`);
			await page.locator('[data-create="xlsx"]').first().click();
			await page.waitForFunction(() => document.querySelector('xlsx-editor')?.workbook);
			await ready(page);
			const ed = page.locator('xlsx-editor');
			for (const [start, col] of [
				['A1', ['Month', 'Jan', 'Feb', 'Mar', 'Apr']],
				['B1', ['Sales', '120', '180', '150', '210']],
			]) {
				await ed.evaluate((e, s) => (e.select(s), e.focusGrid()), start);
				for (const v of col) {
					await page.keyboard.type(v);
					await page.keyboard.press('Enter');
				}
			}
			await ed.evaluate((e) => (e.select('D2'), e.focusGrid()));
			await page.waitForTimeout(800);
			await page.mouse.move(700, 500);
		},
		async play(page) {
			const ed = page.locator('xlsx-editor');
			// A new sheet's default sizes at this viewport: 64 px columns after a 25 px row header,
			// 20 px rows from y = 318 (the grid starts under the ribbon and the formula bar).
			const cell = async (ref) => {
				const [col, row] = [ref.charCodeAt(0) - 65, Number(ref.slice(1))];
				return { x: 25 + 64 * col + 32, y: 318 + 20 * (row - 1) + 10 };
			};
			const a1 = await cell('A1');
			const b5 = await cell('B5');
			await glide(page, a1.x, a1.y);
			await page.mouse.down();
			await page.mouse.move(b5.x, b5.y, { steps: 20 });
			await page.mouse.up();
			await page.waitForTimeout(300);
			await clickOn(page, ed.getByRole('tab', { name: 'Insert' }));
			await page.waitForTimeout(250);
			await clickOn(page, ed.getByRole('button', { name: 'Charts', exact: true }));
			await page.waitForTimeout(250);
			await clickOn(page, ed.getByRole('button', { name: 'Insert Column Chart' }));
			await page.waitForTimeout(500);
			const title = ed.getByRole('textbox', { name: 'Chart title' });
			if (await title.isVisible().catch(() => false)) {
				await clickOn(page, title);
				await type(page, 'Sales by month', 45);
			}
			await page.waitForTimeout(300);
			await clickOn(page, ed.getByRole('button', { name: 'OK', exact: true }));
			await page.waitForTimeout(900);
			const b6 = await cell('B6');
			await glide(page, b6.x, b6.y);
			await page.mouse.down();
			await page.mouse.up();
			await type(page, '=SUM(B2:B5)', 70);
			await page.keyboard.press('Enter');
			await page.waitForTimeout(600);
			const b3 = await cell('B3');
			await glide(page, b3.x, b3.y);
			await page.mouse.down();
			await page.mouse.up();
			await type(page, '260', 90);
			await page.keyboard.press('Enter');
			await page.waitForTimeout(700);
			const b4 = await cell('B4');
			await glide(page, b4.x, b4.y);
			await page.mouse.down();
			await page.mouse.up();
			await type(page, '95', 90);
			await page.keyboard.press('Enter');
			await page.waitForTimeout(1600);
		},
	},
	{
		name: 'powerpoint',
		out: 'viewers/pptx/.github/assets/editor.gif',
		// Dark, graded slides dither heavily: a narrower, slower GIF keeps it near the size budget.
		width: 800,
		fps: 8,
		init: productInit,
		async setup(page, base) {
			await openSample(page, base, '/apps/powerpoint/', 'product-launch-keynote');
			await page.locator('[data-pptx-element="true"]').first().waitFor();
			await page.waitForTimeout(2500);
			await page.mouse.move(600, 500);
		},
		async play(page) {
			await clickOn(page, page.getByRole('button', { name: 'Go to slide 3', exact: true }));
			await page.waitForTimeout(1000);
			// Select the slide title, drag it across the slide, then undo the move.
			const title = page
				.locator('[data-pptx-element="true"]')
				.filter({ hasText: 'Meet Aerie.' })
				.last();
			const target = await center(title);
			await glide(page, target.x, target.y, 600);
			await page.mouse.down();
			await page.mouse.up();
			await page.waitForTimeout(500);
			// Few steps: every move repaints the slide, so a long drag takes seconds in a recording.
			await page.mouse.down();
			await page.mouse.move(target.x + 40, target.y + 10, { steps: 5 });
			await page.mouse.move(target.x + 80, target.y + 25, { steps: 5 });
			await page.mouse.up();
			await page.waitForTimeout(800);
			await page.keyboard.press('Control+z');
			await page.waitForTimeout(800);
			await clickOn(page, page.getByRole('button', { name: 'Go to slide 4', exact: true }));
			await page.waitForTimeout(1100);
			await clickOn(page, page.getByRole('button', { name: 'Go to slide 2', exact: true }));
			await page.waitForTimeout(1100);
		},
	},
	{
		name: 'visio',
		out: 'viewers/visio/docs/public/hero-viewer.gif',
		width: 800,
		init: productInit,
		async setup(page, base) {
			await openSample(page, base, '/apps/visio/', 'order-to-delivery-flowchart');
			await page.waitForFunction(() => Boolean(document.querySelector('visio-viewer')?.document));
			await page.waitForTimeout(2000);
			await page.mouse.move(700, 450);
		},
		async play(page) {
			const viewer = page.locator('visio-viewer');
			const shape = (text) => viewer.locator('g[data-shape-id]').filter({ hasText: text }).last();
			// One click per page: selecting a shape shows its handles and the status bar's size.
			await clickOn(page, shape('Check stock'), { ms: 600 });
			await page.waitForTimeout(900);
			// Zoom in around the pointer, then pan across the drawing.
			await page.keyboard.down('Control');
			for (let i = 0; i < 4; i++) {
				await page.mouse.wheel(0, -120);
				await page.waitForTimeout(120);
			}
			await page.keyboard.up('Control');
			await page.waitForTimeout(500);
			for (let i = 0; i < 6; i++) {
				await page.mouse.wheel(0, 80);
				await page.waitForTimeout(80);
			}
			await page.waitForTimeout(500);
			await clickOn(page, viewer.getByRole('tab', { name: 'Returns' }).first());
			await page.waitForTimeout(1600);
		},
	},
	{
		name: 'teams',
		out: 'viewers/teams/docs/public/hero-workspace.gif',
		init: productInit,
		async setup(page, base) {
			await page.goto(`${base}/apps/teams/`);
			await page.locator('teams-app').waitFor();
			await ready(page);
			await page.waitForTimeout(1500);
			await page.mouse.move(800, 400);
		},
		async play(page) {
			const app = page.locator('teams-app');
			const composer = app.getByRole('textbox').last();
			await clickOn(page, composer);
			await type(page, 'Morning all. The spring range brief is ready for review.', 35);
			await page.keyboard.press('Enter');
			await page.waitForTimeout(700);
			await type(page, 'Can we meet on Thursday at 14:00?', 40);
			await page.keyboard.press('Enter');
			await page.waitForTimeout(600);
			// React to the first message from its hover toolbar.
			const first = app
				.getByText('The spring range brief is ready for review.', { exact: false })
				.first();
			const { x, y } = await center(first);
			await glide(page, x, y, 500);
			await page.waitForTimeout(300);
			await clickOn(page, app.getByRole('button', { name: 'React 👍' }).first());
			await page.waitForTimeout(400);
			await glide(page, x + 200, y + 120, 400);
			await page.waitForTimeout(1200);
		},
	},
];
