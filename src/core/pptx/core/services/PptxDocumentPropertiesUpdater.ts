import type { XMLBuilder, XMLParser } from 'fast-xml-parser';
import type JSZip from 'jszip';

import {
	formatW3cdtf,
	parseAppProperties,
	parseCoreProperties,
	replaceTitleGroup,
	writeAppProperties,
	writeCoreProperties,
	writeCustomProperties,
	type AppProperties,
} from '../../../opc/properties/index';
import type {
	PptxAppProperties,
	PptxCoreProperties,
	PptxCustomProperty,
	PptxSlide,
} from '../types';
import { countNotesPages, toAppTitleEntries } from '../utils/app-properties-counts';
import { countPresentationTextStatistics } from '../utils/app-properties-text-stats';
import {
	applyAppOverrides,
	applyCoreOverrides,
	toCustomProperty,
} from '../utils/document-properties-model';
import { deriveSlideTitles } from '../utils/slide-title';
import {
	ensureCustomPropertiesPackagingArtifacts,
	removeCustomPropertiesPackagingArtifacts,
} from './custom-properties-package';

export interface PptxDocumentPropertiesSaveOptions {
	coreProperties?: PptxCoreProperties;
	appProperties?: PptxAppProperties;
	customProperties?: PptxCustomProperty[];
}

export interface PptxDocumentPropertiesUpdaterContext {
	zip: JSZip;
	parser: XMLParser;
	builder: XMLBuilder;
}

/** The `HeadingPairs` group PowerPoint lists slide titles under. */
const SLIDE_TITLES_GROUP = 'Slide Titles';

const isSlideTitlesGroup = (name: string): boolean => {
	const normalized = name.trim().toLowerCase();
	return (
		normalized === 'slide titles' || (normalized.includes('slide') && normalized.includes('title'))
	);
};

/**
 * Refreshes the document-property parts on save the way PowerPoint does, through the shared
 * `opc/properties` model: `core.xml` gets the next revision and a fresh `dcterms:modified`,
 * `app.xml` the recomputed slide, notes and text statistics and slide titles, and `custom.xml`
 * the caller's custom properties. Each part is patched in place, so everything the save does not
 * change keeps its exact bytes.
 */
export class PptxDocumentPropertiesUpdater {
	private readonly context: PptxDocumentPropertiesUpdaterContext;

	public constructor(context: PptxDocumentPropertiesUpdaterContext) {
		this.context = context;
	}

	public async updateOnSave(
		slides: PptxSlide[],
		options?: PptxDocumentPropertiesSaveOptions,
	): Promise<void> {
		await this.updateCoreProperties(options?.coreProperties);
		await this.updateAppProperties(slides, options?.appProperties);
		await this.updateCustomProperties(options?.customProperties);
	}

	private async updateCoreProperties(overrides: PptxCoreProperties | undefined): Promise<void> {
		const coreFile = this.context.zip.file('docProps/core.xml');
		if (!coreFile) {
			return;
		}
		try {
			const coreXml = await coreFile.async('string');
			const core = applyCoreOverrides(parseCoreProperties(coreXml), overrides);
			const revision = Number.parseInt(core.revision?.trim() ?? '', 10);
			// `cp:lastModifiedBy` is a real author identity, not a value this engine has any basis
			// to invent, so it is only written when the caller overrides it.
			const next = {
				...core,
				revision: String(Number.isFinite(revision) && revision >= 0 ? revision + 1 : 1),
				modified: formatW3cdtf(new Date()),
			};
			this.context.zip.file('docProps/core.xml', writeCoreProperties(next, coreXml));
		} catch (error) {
			console.warn('Failed to update core document properties:', error);
		}
	}

	private async updateAppProperties(
		slides: PptxSlide[],
		overrides: PptxAppProperties | undefined,
	): Promise<void> {
		const appFile = this.context.zip.file('docProps/app.xml');
		if (!appFile) {
			return;
		}
		try {
			const appXml = await appFile.async('string');
			const app: AppProperties = applyAppOverrides(parseAppProperties(appXml), overrides);
			app.slides = slides.length;
			app.hiddenSlides = slides.filter((slide) => slide.hidden).length;
			// PowerPoint counts notes PAGES, empty or not (see `app-properties-counts.ts`); the
			// text-based count is only the fallback for a package with no presentation rels.
			app.notes =
				(await countNotesPages(this.context.zip, this.context.parser)) ??
				slides.filter((slide) => String(slide.notes || '').trim().length > 0).length;
			// PowerPoint recomputes the text statistics too, from slide AND notes text.
			const textStatistics = await countPresentationTextStatistics(
				this.context.zip,
				this.context.parser,
			);
			if (textStatistics) {
				app.words = textStatistics.words;
				app.paragraphs = textStatistics.paragraphs;
			}
			// Without HeadingPairs the title vectors cannot be split safely, so they are left alone.
			if (app.headingPairs) {
				Object.assign(
					app,
					replaceTitleGroup(app.headingPairs, app.titlesOfParts ?? [], {
						name: SLIDE_TITLES_GROUP,
						titles: toAppTitleEntries(deriveSlideTitles(slides)),
						match: isSlideTitlesGroup,
					}),
				);
			}
			this.context.zip.file('docProps/app.xml', writeAppProperties(app, appXml));
		} catch (error) {
			console.warn('Failed to update application document properties:', error);
		}
	}

	private async updateCustomProperties(
		customProperties: PptxCustomProperty[] | undefined,
	): Promise<void> {
		if (!customProperties) {
			return;
		}
		const customXml = writeCustomProperties(
			customProperties.filter((entry) => entry.name.trim().length > 0).map(toCustomProperty),
		);
		if (customXml === undefined) {
			this.context.zip.remove('docProps/custom.xml');
			await removeCustomPropertiesPackagingArtifacts(this.context);
			return;
		}
		this.context.zip.file('docProps/custom.xml', customXml);
		await ensureCustomPropertiesPackagingArtifacts(this.context);
	}
}
