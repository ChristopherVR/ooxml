// The header and footer bands of a page in Page Layout view: rendered text with the field codes
// (&P, &N, &D, &T, &F, &A) resolved, the "Add header" placeholder, and the inline editor that
// saves the three sections back to the sheet's page setup.
import {
	composeHeaderFooter,
	headerFooterKey,
	pageHeaderFooter,
	renderHeaderFooterRuns,
	splitHeaderFooter,
	type HeaderFooterKind,
	type HeaderFooterSection,
	type PageViewLayout,
	type PageViewPage,
} from 'ooxml-core/xlsx';
import type { EditorContext } from 'ooxml-core/xlsx/ui';

const SECTIONS: readonly HeaderFooterSection[] = ['left', 'center', 'right'];
const SECTION_LABELS = {
	left: 'Left section',
	center: 'Center section',
	right: 'Right section',
} as const;

export interface BandOptions {
	ctx: EditorContext;
	doc: Document;
	layout: PageViewLayout;
	sheetIndex: number;
	fileName: string;
	date: Date;
}

const pageNumber = (o: BandOptions, page: PageViewPage): number =>
	(o.ctx.workbook()?.sheets[o.sheetIndex]?.pageSetup?.firstPageNumber ?? 1) + page.number - 1;

/** The band element for a page's header or footer, positioned inside the page. */
export function renderBand(
	o: BandOptions,
	page: PageViewPage,
	kind: HeaderFooterKind,
): HTMLElement {
	const sheet = o.ctx.workbook()!.sheets[o.sheetIndex]!;
	const setup = sheet.pageSetup ?? {};
	const number = pageNumber(o, page);
	const value = pageHeaderFooter(setup, page.number, number)[kind];
	const rect = page[kind];
	const band = o.doc.createElement('div');
	band.className = `xpl-band xpl-${kind}`;
	band.dataset.kind = kind;
	band.style.cssText = `left:${rect.x - page.paper.x}px;top:${rect.y - page.paper.y}px;width:${rect.width}px;min-height:${Math.max(rect.height, 18)}px`;
	const size = ((11 * 96) / 72) * (setup.scaleWithDoc === false ? 1 : o.layout.scale);
	band.style.fontSize = `${size}px`;
	const rendered = renderHeaderFooterRuns(value, {
		page: number,
		pages: o.layout.pages.length,
		sheet: sheet.name,
		file: o.fileName,
		date: o.date,
		locale: o.ctx.locale(),
	});
	const empty = !value;
	for (const section of SECTIONS) {
		const cell = o.doc.createElement('span');
		cell.className = 'xpl-section';
		cell.dataset.section = section;
		cell.style.textAlign = section;
		const runs = rendered.sections[section];
		for (const run of runs) {
			const span = o.doc.createElement('span');
			span.textContent = run.text;
			const { style } = run;
			if (style.font) span.style.fontFamily = style.font;
			if (style.size) span.style.fontSize = `${(style.size / 11) * size}px`;
			if (style.color) span.style.color = style.color;
			if (style.bold) span.style.fontWeight = '700';
			if (style.italic) span.style.fontStyle = 'italic';
			const lines = [style.underline && 'underline', style.strike && 'line-through'].filter(
				Boolean,
			);
			if (lines.length) span.style.textDecoration = lines.join(' ');
			cell.append(span);
		}
		if (empty && section === 'center') {
			cell.classList.add('xpl-placeholder');
			cell.textContent = o.ctx.t(kind === 'header' ? 'Add header' : 'Add footer');
		}
		band.append(cell);
	}
	return band;
}

/**
 * Turns a band into three text boxes (the raw text, field codes visible) focused on `section`.
 * Leaving the band, or Enter, saves the sections to the page setup; Escape discards the edit.
 */
export function editBand(
	o: BandOptions,
	page: PageViewPage,
	band: HTMLElement,
	kind: HeaderFooterKind,
	section: HeaderFooterSection,
	done: () => void,
): void {
	const session = o.ctx.session();
	const sheet = o.ctx.workbook()?.sheets[o.sheetIndex];
	if (!session || !sheet || o.ctx.readOnly()) return;
	const setup = sheet.pageSetup ?? {};
	const key = headerFooterKey(setup, kind, page.number, pageNumber(o, page));
	const current = splitHeaderFooter(setup[key] ?? '');
	band.classList.add('xpl-editing');
	band.replaceChildren();
	const inputs = {} as Record<HeaderFooterSection, HTMLInputElement>;
	for (const name of SECTIONS) {
		const input = o.doc.createElement('input');
		input.type = 'text';
		input.className = 'xpl-band-input';
		input.value = current[name];
		input.spellcheck = false;
		input.dataset.section = name;
		input.setAttribute('aria-label', o.ctx.t(SECTION_LABELS[name]));
		input.title = o.ctx.t('Codes: &P page, &N pages, &D date, &T time, &F file, &A sheet');
		band.append(input);
		inputs[name] = input;
	}
	let finished = false;
	const finish = (save: boolean) => {
		if (finished) return;
		finished = true;
		if (save) {
			const text = composeHeaderFooter({
				left: inputs.left.value,
				center: inputs.center.value,
				right: inputs.right.value,
			});
			if (text !== (setup[key] ?? ''))
				session.setPageSetup(o.sheetIndex, { [key]: text || undefined });
		}
		done();
	};
	band.addEventListener('focusout', (event) => {
		const next = event.relatedTarget as Node | null;
		if (!next || !band.contains(next)) finish(true);
	});
	band.addEventListener('keydown', (event) => {
		if (event.key === 'Enter') {
			event.preventDefault();
			finish(true);
		} else if (event.key === 'Escape') {
			event.preventDefault();
			finish(false);
		}
		event.stopPropagation();
	});
	inputs[section].focus();
	inputs[section].select();
}
