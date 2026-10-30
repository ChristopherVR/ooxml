// @vitest-environment jsdom
import { afterEach, expect, it, vi } from 'vitest';
import { demoDocument } from 'ooxml-core/visio/ui';
import { mountViewer } from './binding';
import { localizeWording } from './ribbon-localize';
import { pdfOptionsPanel, showPdfReport, type PdfReportSummary } from './viewer-pdf-options';

afterEach(() => {
	vi.restoreAllMocks();
	document.body.replaceChildren();
});

it('translates the PDF options panel and the export report with the locale', () => {
	const panel = pdfOptionsPanel(document);
	document.body.append(panel);
	const summary = {
		pages: [
			{ name: 'Page-1', mode: 'vector', fallback: [], degraded: [] },
			{ name: 'Page-2', mode: 'picture', fallback: [], degraded: [] },
		],
		embeddedFonts: ['Arial'],
		refusedFonts: [],
	} as unknown as PdfReportSummary;
	showPdfReport(panel, summary);
	const walk = (locale: 'en' | 'de') =>
		localizeWording([panel, ...panel.querySelectorAll<HTMLElement>('*')], locale);
	const lines = () =>
		[...panel.querySelectorAll('[data-pdf-report] li, [data-pdf-report] strong')].map(
			(node) => node.textContent,
		);

	walk('en');
	expect(panel.querySelector('h2')!.textContent).toBe('PDF options');
	expect(lines()).toEqual([
		'PDF created: 2 page(s)',
		'Page-1: drawn as vectors with selectable text.',
		'Page-2: saved as a picture (it cannot be drawn as vectors).',
		'Fonts embedded: Arial.',
	]);

	walk('de');
	expect(panel.querySelector('h2')!.textContent).toBe('PDF-Optionen');
	expect(lines()).toEqual([
		'PDF erstellt: 2 Seite(n)',
		'Page-1: als Vektoren mit auswählbarem Text gezeichnet.',
		'Page-2: als Bild gespeichert (sie lässt sich nicht als Vektoren zeichnen).',
		'Eingebettete Schriftarten: Arial.',
	]);
	panel.remove();
});

it('translates the File backstage when the locale changes, and back to English', () => {
	const host = document.createElement('div');
	document.body.append(host);
	const viewer = mountViewer(host, { document: demoDocument });
	const element = viewer.element;
	const root = element.shadowRoot!;
	const backstage = root.querySelector<HTMLElement & { items: { id: string; label: string }[] }>(
		'office-ui-backstage',
	)!;
	const text = (selector: string) => root.querySelector(selector)!.textContent;
	const navLabel = (id: string) => backstage.items.find((item) => item.id === id)?.label;

	element.setAttribute('locale', 'en');
	element.openBackstage('info');
	expect(text('[data-backstage-page="info"] h1')).toBe('Info');
	expect(navLabel('info')).toBe('Info');
	expect(text('[data-backstage-action="download"] strong')).toBe('Download a copy');

	element.setAttribute('locale', 'de');
	expect(text('[data-backstage-page="info"] h1')).toBe('Informationen');
	expect(navLabel('info')).toBe('Informationen');
	expect(backstage.getAttribute('back-label')).toBe('Zurück zur Zeichnung');
	expect(text('[data-backstage-page="info"] .backstage-note-count')).toBe(
		'Ein Teil dieser Zeichnung wird nur ungefähr oder gar nicht angezeigt:',
	);
	// The Export page's PDF panel and the tiles of the New page follow the locale too.
	expect(text('[data-backstage-page="export"] .pdf-options h2')).toBe('PDF-Optionen');
	expect(text('[data-backstage-action="new-blank"] strong')).toBe('Leere Zeichnung');
	expect(text('[data-backstage-action="export-pdf"] strong')).toBe('PDF/XPS-Dokument erstellen');

	element.setAttribute('locale', 'en');
	expect(text('[data-backstage-page="info"] h1')).toBe('Info');
	expect(navLabel('info')).toBe('Info');
	expect(backstage.getAttribute('back-label')).toBe('Back to the drawing');
	expect(text('[data-backstage-page="info"] .backstage-note-count')).toBe(
		'One part of this drawing is shown approximately or not at all:',
	);
	expect(text('[data-backstage-action="new-blank"] strong')).toBe('Blank Drawing');
	viewer.destroy();
});
