import { twips, twipsToPixels, type Twips } from '../units/units';
import type { DocumentModel } from './model';
import { sectionsOf } from './section-layout';
import type { SectionProperties } from './section-model';

const px = twipsToPixels;

/** Returns `model` with section `index` changed; the last section also drives `model.page`. */
export function withSection(
	model: DocumentModel,
	index: number,
	change: (section: SectionProperties) => SectionProperties,
): DocumentModel {
	const sections = sectionsOf(model).map((section, position) =>
		position === index ? change(structuredClone(section)) : section,
	);
	const last = sections.at(-1)!;
	return {
		...model,
		sections,
		page: {
			width: px(last.pageWidthTwips),
			height: px(last.pageHeightTwips),
			marginTop: px(last.marginTopTwips),
			marginRight: px(last.marginRightTwips),
			marginBottom: px(last.marginBottomTwips),
			marginLeft: px(last.marginLeftTwips),
		},
	};
}

const NORMAL_MARGIN = twips(1440);
const MARGINS: Record<string, Twips> = {
	normal: NORMAL_MARGIN,
	narrow: twips(720),
	wide: twips(2160),
};

export function setMargins(model: DocumentModel, index: number, preset: string): DocumentModel {
	if (preset === 'moderate')
		return withSection(model, index, (section) => ({
			...section,
			marginTopTwips: NORMAL_MARGIN,
			marginBottomTwips: NORMAL_MARGIN,
			marginLeftTwips: twips(1080),
			marginRightTwips: twips(1080),
		}));
	const value = MARGINS[preset] ?? NORMAL_MARGIN;
	return withSection(model, index, (section) => ({
		...section,
		marginTopTwips: value,
		marginRightTwips: value,
		marginBottomTwips: value,
		marginLeftTwips: value,
	}));
}

export function setOrientation(
	model: DocumentModel,
	index: number,
	orientation: 'portrait' | 'landscape',
): DocumentModel {
	return withSection(model, index, (section) => {
		const { pageWidthTwips: width, pageHeightTwips: height } = section;
		const long = width >= height ? width : height;
		const short = width >= height ? height : width;
		return {
			...section,
			orientation,
			pageWidthTwips: orientation === 'landscape' ? long : short,
			pageHeightTwips: orientation === 'landscape' ? short : long,
		};
	});
}

export function setColumns(model: DocumentModel, index: number, count: number): DocumentModel {
	return withSection(model, index, (section) => ({
		...section,
		columns: { count, spacingTwips: section.columns.spacingTwips ?? twips(720), equalWidth: true },
	}));
}

/** Layout > Page Setup > Vertical alignment for one section. */
export function setVerticalAlign(
	model: DocumentModel,
	index: number,
	verticalAlign: NonNullable<SectionProperties['verticalAlign']>,
): DocumentModel {
	return withSection(model, index, (section) => {
		const { verticalAlign: _previous, ...rest } = section;
		return verticalAlign === 'top' ? rest : { ...rest, verticalAlign };
	});
}

/** Header & Footer > Different First Page for one section. */
export function setTitlePage(
	model: DocumentModel,
	index: number,
	titlePage: boolean,
): DocumentModel {
	return withSection(model, index, (section) => ({ ...section, titlePage }));
}

/** Page number format, and whether numbering continues or restarts at 1 in this section. */
export function setPageNumbering(
	model: DocumentModel,
	index: number,
	change: { format?: NonNullable<SectionProperties['pageNumbering']>['format']; restart?: boolean },
): DocumentModel {
	return withSection(model, index, (section) => {
		const numbering = { ...section.pageNumbering };
		if (change.format !== undefined) numbering.format = change.format;
		if (change.restart === true) numbering.start = 1;
		if (change.restart === false) delete numbering.start;
		if (numbering.format === 'decimal') delete numbering.format;
		const { pageNumbering: _previous, ...rest } = section;
		return Object.keys(numbering).length ? { ...rest, pageNumbering: numbering } : rest;
	});
}

/** Layout > Line Numbers: off, or on with the restart rule; other line-number values are kept. */
export function setLineNumbering(
	model: DocumentModel,
	index: number,
	mode: 'none' | 'continuous' | 'newPage' | 'newSection',
): DocumentModel {
	return withSection(model, index, (section) => {
		const { lineNumbering: _on, lineNumberSettings: previous, ...rest } = section;
		if (mode === 'none') return rest;
		return {
			...rest,
			lineNumbering: true,
			lineNumberSettings: {
				countBy: previous?.countBy ?? 1,
				start: previous?.start ?? 1,
				restart: mode,
				...(previous?.distanceTwips !== undefined ? { distanceTwips: previous.distanceTwips } : {}),
			},
		};
	});
}
