// @vitest-environment node
/**
 * Server rendering inside a `PptxI18nProvider` with no i18next instance anywhere: the viewer
 * and a standalone building block render the host's strings and never a raw key.
 */
import { renderToString } from 'react-dom/server';
import { describe, expect, it } from 'vitest';

import { PowerPointViewer, PptxI18nProvider, SlideTemplateGalleryDialog } from '../../index';
import type { PptxTranslate } from '../../index';

const HOST: Readonly<Record<string, string>> = {
	'pptx.viewer.loading': 'Laedt (host)',
	'pptx.slideTemplates.galleryTitle': 'Vorlagen (host)',
};
const translate: PptxTranslate = (key) => HOST[key];

/** Visible text of server markup, without tags. */
const textOf = (html: string) => html.replace(/<[^>]*>/g, ' ');

describe('pptxI18nProvider on the server', () => {
	it('renders the viewer from the host translator', () => {
		const html = renderToString(
			<PptxI18nProvider translate={translate} locale='de'>
				<PowerPointViewer content={null} canEdit />
			</PptxI18nProvider>,
		);
		expect(html).toContain('Laedt (host)');
		expect(textOf(html)).not.toMatch(/\bpptx\.[A-Za-z]/);
	});

	it('renders a standalone building block from the host translator', () => {
		const html = renderToString(
			<PptxI18nProvider translate={translate} locale='de'>
				<SlideTemplateGalleryDialog isOpen onClose={() => {}} onInsert={() => {}} />
			</PptxI18nProvider>,
		);
		expect(html).toContain('Vorlagen (host)');
		expect(textOf(html)).not.toMatch(/\bpptx\.[A-Za-z]/);
	});
});
