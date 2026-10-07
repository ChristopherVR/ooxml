import { NS, children, first, type XmlElement } from '../xml';
import type { ChartStyleDefinition, ChartStyleEntry, ChartStylePart } from './style-definition';

const C14 = NS.c14;

/** Newly authored Office charts use modern style 102; imports retain their parsed choice. */
export function effectiveBuiltInChartStyle(
	formatting: ChartStyleDefinition | undefined,
): number | undefined {
	return formatting?.builtInStyle ?? (formatting?.sourceXml ? undefined : 102);
}

/** Prefer the supported Office 2010 chart style over its legacy fallback. */
export function readBuiltInChartStyle(root: XmlElement): number | undefined {
	for (const content of children(root, 'AlternateContent', NS.mc)) {
		for (const choice of children(content, 'Choice', NS.mc)) {
			const requires = choice.getAttribute('Requires')?.trim().split(/\s+/);
			if (
				!requires?.length ||
				!requires.every((prefix) => choice.lookupNamespaceURI(prefix) === C14)
			)
				continue;
			const modern = first(choice, 'style', C14)?.getAttribute('val');
			if (modern && /^\d+$/.test(modern) && Number(modern) >= 101 && Number(modern) <= 148)
				return Number(modern);
		}
	}
	const legacy =
		first(root, 'style', NS.c) ??
		first(first(first(root, 'AlternateContent', NS.mc), 'Fallback', NS.mc), 'style', NS.c);
	const value = legacy?.getAttribute('val');
	if (!legacy) return 2;
	return value && /^\d+$/.test(value) && Number.isSafeInteger(Number(value))
		? Number(value)
		: undefined;
}

/** Preserve known built-in defaults when a chart part has to be regenerated. */
export function builtInChartStyleXml(id: number | undefined): string {
	if (id === undefined || !Number.isInteger(id)) return '';
	if (id >= 1 && id <= 48) return `<c:style val="${id}"/>`;
	if (id < 101 || id > 148) return '';
	return `<mc:AlternateContent xmlns:mc="${NS.mc}"><mc:Choice Requires="c14" xmlns:c14="${C14}"><c14:style val="${id}"/></mc:Choice><mc:Fallback><c:style val="${id - 100}"/></mc:Fallback></mc:AlternateContent>`;
}

/** COM-verified built-in text defaults; values stay in native points and theme choices. */
export function builtInChartTextStyle(
	id: number | undefined,
): Partial<Record<ChartStylePart, ChartStyleEntry>> {
	if (
		id === undefined ||
		!Number.isInteger(id) ||
		!((id >= 1 && id <= 48) || (id >= 101 && id <= 148))
	)
		return {};
	const legacy = id > 100 ? id - 100 : id;
	const common: ChartStyleEntry = {
		sourceXml: '',
		fontSize: 10,
		bold: false,
		italic: false,
		fontRef: {
			index: 'minor',
			color: { kind: 'scheme', value: legacy >= 41 ? 'lt1' : 'tx1', transforms: [] },
		},
	};
	return {
		title: { ...common, fontSize: id > 100 ? 18 : 10, bold: true },
		categoryAxis: { ...common },
		valueAxis: { ...common },
		legend: { ...common },
	};
}
