/**
 * A deck under construction: the hand-written template is loaded by the core's `PptxHandler`,
 * slides are added through the SDK, and the handler's own `save` writes the package. What the
 * save cannot produce yet (paragraph builds) is patched afterwards in `finish.ts`.
 */
import { PptxHandler } from '../../../../src/core/pptx/core/PptxHandler';
import type { PptxData, PptxSection } from '../../../../src/core/pptx/core/types/presentation';
import { seededIdSource } from '../../../../src/core/opc/reproducible';
import { buildPackage } from '../../lib/package';
import { SAMPLE_DATE, SAMPLE_ID_SEED } from '../../lib/types';
import { finishPackage } from './finish';
import { Slide } from './slide';
import { buildTemplate, type TemplateSpec } from './template';

export class Deck {
	readonly slides: Slide[] = [];
	private readonly sections: PptxSection[] = [];

	private constructor(
		private readonly spec: TemplateSpec,
		private readonly handler: PptxHandler,
		private readonly data: PptxData,
	) {}

	static async open(spec: TemplateSpec): Promise<Deck> {
		const bytes = await buildPackage(buildTemplate(spec));
		const handler = new PptxHandler();
		const copy = new Uint8Array(bytes);
		const data = await handler.load(copy.buffer);
		return new Deck(spec, handler, data);
	}

	/** Starts a named section; the slides added next belong to it. */
	section(name: string): this {
		const serial = String(this.sections.length + 1).padStart(12, '0');
		this.sections.push({ id: `{6F1D2C3A-4B5E-4A6F-8C7D-${serial}}`, name, slideIds: [] });
		return this;
	}

	/** Adds a slide on the named layout of the template. */
	slide(layoutName: string): Slide {
		const index = this.spec.layouts.findIndex((layout) => layout.name === layoutName);
		if (index < 0) throw new Error(`The template has no layout named "${layoutName}".`);
		const slide = new Slide(
			this.slides.length + 1,
			`ppt/slideLayouts/slideLayout${index + 1}.xml`,
			layoutName,
			this.handler,
		);
		const section = this.sections[this.sections.length - 1];
		if (section) {
			slide.model.sectionId = section.id;
			slide.model.sectionName = section.name;
			section.slideIds.push(slide.model.id);
		}
		this.slides.push(slide);
		return slide;
	}

	async save(): Promise<Uint8Array> {
		const models = this.slides.map((slide) => slide.model);
		this.data.slides.push(...models);
		// A reproducible save: the sample date and ids from a seeded source, so two builds agree.
		const saved = await this.handler.save(this.data.slides, {
			sections: this.sections,
			fixedDate: SAMPLE_DATE,
			ids: seededIdSource(SAMPLE_ID_SEED),
		});
		const warnings = this.handler.getCompatibilityWarnings();
		if (warnings.length > 0)
			throw new Error(`The core reported save warnings: ${JSON.stringify(warnings)}`);
		return finishPackage(saved, this.slides);
	}
}
