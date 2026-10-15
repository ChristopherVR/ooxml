import JSZip from 'jszip';
import { afterEach, expect, it, vi } from 'vitest';
import { createVsdx, editVsdx, parseVsdx } from 'ooxml-core/visio';
import { ViewerController } from './controller';
import { createShapesWindow } from './shapes-window';
import { wireStencil } from './viewer-stencil';

afterEach(() => {
	document.body.replaceChildren();
	localStorage.clear();
});

const NS = 'http://schemas.microsoft.com/office/visio/2012/main';
const RELS = 'http://schemas.openxmlformats.org/package/2006/relationships';
const REL = 'http://schemas.microsoft.com/visio/2010/relationships/';
const cell = (name: string, value: number | string, formula = '') =>
	`<Cell N="${name}" V="${value}"${formula ? ` F="${formula}"` : ''}/>`;
const row = (index: number, type: string, cells: string) =>
	`<Row IX="${index}" T="${type}">${cells}</Row>`;

/** A .vssx with one box master that has a connection point and a Shape Data row. */
async function vssx(): Promise<Uint8Array> {
	const zip = new JSZip();
	zip.file(
		'[Content_Types].xml',
		'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Override PartName="/visio/document.xml" ContentType="application/vnd.ms-visio.stencil.main+xml"/></Types>',
	);
	zip.file(
		'_rels/.rels',
		`<Relationships xmlns="${RELS}"><Relationship Id="rId1" Type="${REL}document" Target="visio/document.xml"/></Relationships>`,
	);
	zip.file('visio/document.xml', `<VisioDocument xmlns="${NS}"/>`);
	zip.file(
		'visio/_rels/document.xml.rels',
		`<Relationships xmlns="${RELS}"><Relationship Id="rId1" Type="${REL}masters" Target="masters/masters.xml"/></Relationships>`,
	);
	zip.file(
		'visio/masters/masters.xml',
		`<Masters xmlns="${NS}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><Master ID="2" NameU="Valve" Name="Valve" UniqueID="{AAAAAAAA-0000-0000-0000-000000000001}"><PageSheet>${cell('PageWidth', 1)}${cell('PageHeight', 1)}</PageSheet><Rel r:id="rId1"/></Master></Masters>`,
	);
	zip.file(
		'visio/masters/_rels/masters.xml.rels',
		`<Relationships xmlns="${RELS}"><Relationship Id="rId1" Type="${REL}master" Target="master1.xml"/></Relationships>`,
	);
	const geometry = `<Section N="Geometry" IX="0">${row(1, 'RelMoveTo', cell('X', 0) + cell('Y', 0))}${row(2, 'RelLineTo', cell('X', 1) + cell('Y', 0))}${row(3, 'RelLineTo', cell('X', 1) + cell('Y', 1))}${row(4, 'RelLineTo', cell('X', 0) + cell('Y', 1))}${row(5, 'RelLineTo', cell('X', 0) + cell('Y', 0))}</Section>`;
	const points = `<Section N="Connection">${row(0, 'Connection', cell('X', 0.5, 'Width*0.5') + cell('Y', 1, 'Height*1'))}</Section>`;
	const data = `<Section N="Property"><Row N="Pressure">${cell('Label', 'Pressure')}${cell('Value', 'bar')}${cell('Type', 0)}</Row></Section>`;
	zip.file(
		'visio/masters/master1.xml',
		`<MasterContents xmlns="${NS}"><Shapes><Shape ID="1" Type="Shape">${cell('PinX', 0.5)}${cell('PinY', 0.5)}${cell('Width', 1)}${cell('Height', 1)}${cell('LocPinX', 0.5, 'Width*0.5')}${cell('LocPinY', 0.5, 'Height*0.5')}${points}${data}${geometry}</Shape></Shapes></MasterContents>`,
	);
	return zip.generateAsync({ type: 'uint8array' });
}

async function open() {
	const controller = new ViewerController(
		parseVsdx,
		vi.fn(),
		async (bytes, commands) => {
			const result = await editVsdx(bytes, commands);
			return { ...result, document: await parseVsdx(result.bytes) };
		},
		async () => 'unused',
	);
	const pane = createShapesWindow(document);
	const viewport = document.createElement('div');
	document.body.append(pane, viewport);
	const messages: string[] = [];
	const dispose = wireStencil(pane, viewport, controller, (message) => messages.push(message));
	await controller.load(await createVsdx());
	const choose = async (name: string, bytes: Uint8Array) => {
		const input = pane.querySelector<HTMLInputElement>('input[data-open-stencil]')!;
		const file = new File([bytes as BlobPart], name);
		Object.defineProperty(input, 'files', { value: [file], configurable: true });
		input.dispatchEvent(new Event('change'));
	};
	return { controller, pane, messages, choose, dispose };
}

it('opens a .vssx from More Shapes, drops its master and saves an instance of it', async () => {
	const ui = await open();
	const item = ui.pane.querySelector('[data-shapes-menu="more"] [command="open-stencil-file"]')!;
	expect(item.hasAttribute('disabled')).toBe(false);
	await ui.choose('Valves.vssx', await vssx());
	await vi.waitFor(() => expect(ui.pane.querySelector('[data-stencil="file-1"]')).not.toBeNull());
	const section = ui.pane.querySelector<HTMLElement>('[data-stencil="file-1"]')!;
	expect(section.querySelector('.stencil-title')!.textContent).toBe('Valves');
	const button = section.querySelector<HTMLButtonElement>('[data-master]')!;
	expect(button.dataset.master).toBe('file:1:2');
	button.click();
	await vi.waitFor(() => expect(ui.messages.length).toBe(1), { timeout: 5000 });
	const model = ui.controller.state.document!;
	expect(model.masters!.map((master) => master.name)).toEqual(['Valve']);
	const dropped = model.pages[0]!.shapes[0]!;
	expect(dropped.masterId).toBe(model.masters![0]!.id);
	expect(dropped.connectionPoints).toHaveLength(1);
	expect(dropped.shapeData?.map((data) => data.label)).toEqual(['Pressure']);
	ui.dispose();
});

it('collects masters into My Shapes for the session', async () => {
	const ui = await open();
	await ui.choose('Valves.vssx', await vssx());
	await vi.waitFor(() => expect(ui.pane.querySelector('[data-stencil="file-1"]')).not.toBeNull());
	const button = ui.pane.querySelector<HTMLElement>('[data-stencil="file-1"] [data-master]')!;
	button.dispatchEvent(new MouseEvent('contextmenu', { bubbles: true, cancelable: true }));
	ui.pane
		.querySelector('[data-shapes-menu="master"]')!
		.dispatchEvent(
			new CustomEvent('office-command', { detail: { command: 'my-shapes' }, bubbles: true }),
		);
	const mine = ui.pane.querySelector<HTMLElement>('[data-stencil="my-shapes"]')!;
	expect(mine.querySelector('.stencil-title')!.textContent).toBe('My Shapes');
	expect(mine.querySelectorAll('[data-master]')).toHaveLength(1);
	// Closing the stencil keeps what was collected.
	ui.pane.querySelector<HTMLElement>('[data-close-stencil="file-1"]')!.click();
	expect(ui.pane.querySelector('[data-stencil="file-1"]')).toBeNull();
	expect(ui.pane.querySelector('[data-stencil="my-shapes"] [data-master]')).not.toBeNull();
	ui.dispose();
});

it('says why a legacy .vss or a damaged file is not opened', async () => {
	const ui = await open();
	await ui.choose('Old.vss', new Uint8Array([0xd0, 0xcf, 0x11, 0xe0]));
	const status = ui.pane.querySelector<HTMLElement>('[data-stencil-status]')!;
	await vi.waitFor(() => expect(status.hidden).toBe(false));
	expect(status.textContent).toContain('not supported');
	await ui.choose('Broken.vssx', new Uint8Array([1, 2, 3]));
	await vi.waitFor(() => expect(status.textContent).toContain('Broken.vssx'));
	expect(ui.pane.querySelector('[data-stencil^="file-"]')).toBeNull();
	ui.dispose();
});
