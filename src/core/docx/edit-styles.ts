import type {
	DocumentModel,
	ParagraphFormatting,
	RunFormatting,
	TableStyleDefinition,
} from './model';
import { defaultStyleId } from './style-inheritance';
import type { StyleMetadata } from './style-metadata';

export type StyleCatalogs = Pick<
	DocumentModel,
	'paragraphStyles' | 'characterStyles' | 'tableStyles'
>;
export type StyleKind = 'paragraph' | 'character' | 'linked' | 'table';

interface StyleEdits {
	id?: string;
	name: string;
	basedOn?: string;
	uiPriority?: number;
	quickFormat?: boolean;
	runFormatting: RunFormatting;
}

type TableProperties = Pick<
	TableStyleDefinition,
	| 'borders'
	| 'cellMargins'
	| 'conditional'
	| 'shadingFill'
	| 'shadingThemeFill'
	| 'percentageShading'
>;

export type StyleEdit = StyleEdits &
	(
		| { kind: 'character' }
		| {
				kind: 'table';
				paragraphFormatting: ParagraphFormatting;
				tableProperties?: Partial<TableProperties>;
		  }
		| { kind: 'paragraph' | 'linked'; next?: string; paragraphFormatting: ParagraphFormatting }
	);

export function paragraphStyleId(
	catalogs: StyleCatalogs,
	explicit: string | undefined,
): string | undefined {
	return explicit || (catalogs.paragraphStyles && defaultStyleId(catalogs.paragraphStyles.styles));
}

export function styleCatalogsOf(model: DocumentModel): StyleCatalogs {
	return {
		...(model.paragraphStyles && { paragraphStyles: model.paragraphStyles }),
		...(model.characterStyles && { characterStyles: model.characterStyles }),
		...(model.tableStyles && { tableStyles: model.tableStyles }),
	};
}

export function stylesOf(catalogs: StyleCatalogs, kind: StyleKind): Record<string, StyleMetadata> {
	if (kind === 'character')
		return Object.fromEntries(
			Object.entries(catalogs.characterStyles?.styles ?? {}).filter(
				([, style]) => style.type === 'character',
			),
		);
	if (kind === 'table') return catalogs.tableStyles?.styles ?? {};
	return catalogs.paragraphStyles?.styles ?? {};
}

const ownStyle = <T>(styles: Record<string, T>, id: string): T | undefined =>
	Object.hasOwn(styles, id) ? styles[id] : undefined;

export function canBaseStyleOn(
	styles: Record<string, StyleMetadata>,
	base: string,
	id?: string,
): boolean {
	const seen = new Set<string>();
	let current: string | undefined = base;
	while (current) {
		if (current === id || seen.has(current)) return false;
		seen.add(current);
		current = ownStyle(styles, current)?.basedOn;
	}
	return true;
}

function availableId(catalogs: StyleCatalogs, seed: string): string {
	const occupied = (id: string) =>
		ownStyle(catalogs.paragraphStyles?.styles ?? {}, id) ||
		ownStyle(catalogs.characterStyles?.styles ?? {}, id) ||
		ownStyle(catalogs.tableStyles?.styles ?? {}, id) ||
		catalogs.characterStyles?.reservedStyleIds?.includes(id);
	let id = seed;
	for (let suffix = 2; occupied(id); suffix++) id = seed + suffix;
	return id;
}

/** Creates a style or applies edits to its current definition, retaining its id and related data. */
export function editStyleCatalogs(
	catalogs: StyleCatalogs,
	edit: StyleEdit,
): { catalogs: StyleCatalogs; id: string; linkedId?: string } | undefined {
	const name = edit.name.trim();
	const styles = stylesOf(catalogs, edit.kind);
	const old = edit.id === undefined ? undefined : ownStyle(styles, edit.id);
	if (!name || (edit.id !== undefined && !old)) return undefined;
	if (
		edit.basedOn &&
		edit.basedOn !== old?.basedOn &&
		(!ownStyle(styles, edit.basedOn) || !canBaseStyleOn(styles, edit.basedOn, edit.id))
	)
		return undefined;
	const id =
		edit.id ?? availableId(catalogs, name.replace(/[^\p{L}\p{N}_]/gu, '') || 'CustomStyle');
	const priorParagraph = ownStyle(catalogs.paragraphStyles?.styles ?? {}, id);
	const priorRun = ownStyle(catalogs.characterStyles?.styles ?? {}, id);
	const next = { ...catalogs };
	const uiPriority = edit.uiPriority ?? old?.uiPriority;
	const quickFormat = edit.quickFormat ?? old?.quickFormat ?? (old ? undefined : true);
	const metadata = {
		id,
		name,
		...(edit.basedOn && { basedOn: edit.basedOn }),
		...(old?.isDefault !== undefined && { isDefault: old.isDefault }),
		...(old?.customStyle !== undefined
			? { customStyle: old.customStyle }
			: !old && { customStyle: true }),
		...(uiPriority !== undefined && { uiPriority }),
		...(quickFormat !== undefined && { quickFormat }),
	};
	if (edit.kind === 'table') {
		const table = catalogs.tableStyles ?? { styles: {}, warnings: [] };
		const { basedOn: _base, ...properties } = ownStyle(table.styles, id) ?? { conditional: {} };
		const definition: TableStyleDefinition = {
			...properties,
			...edit.tableProperties,
			...metadata,
			runFormatting: edit.runFormatting,
			paragraphFormatting: edit.paragraphFormatting,
		};
		next.tableStyles = { ...table, styles: { ...table.styles, [id]: definition } };
		return { catalogs: next, id };
	}
	const run = catalogs.characterStyles ?? { styles: {}, warnings: [], docDefaults: {} };
	const runStyles = { ...run.styles };
	const oldLink = priorParagraph?.linkedStyle ?? priorRun?.linkedStyle;
	const partner = oldLink && ownStyle(run.styles, oldLink);
	const existingPartner =
		partner && partner.type === 'character' && partner.linkedStyle === id ? partner : undefined;
	const linkedId =
		edit.kind === 'linked' ? (existingPartner?.id ?? availableId(catalogs, id + 'Char')) : oldLink;
	runStyles[id] = {
		...metadata,
		type: edit.kind === 'character' ? 'character' : 'paragraph',
		formatting: edit.runFormatting,
		...(linkedId && { linkedStyle: linkedId }),
		...(priorRun?.eastAsiaLanguage !== undefined && {
			eastAsiaLanguage: priorRun.eastAsiaLanguage,
		}),
	};
	next.characterStyles = { ...run, styles: runStyles };
	if (edit.kind !== 'character') {
		const paragraph = catalogs.paragraphStyles ?? { styles: {}, warnings: [], docDefaults: {} };
		if (
			edit.next &&
			edit.next !== id &&
			edit.next !== priorParagraph?.next &&
			!ownStyle(paragraph.styles, edit.next)
		)
			return undefined;
		next.paragraphStyles = {
			...paragraph,
			styles: {
				...paragraph.styles,
				[id]: {
					...metadata,
					...(linkedId && { linkedStyle: linkedId }),
					...(edit.next && { next: edit.next }),
					...(priorParagraph?.numbering && { numbering: priorParagraph.numbering }),
					formatting: edit.paragraphFormatting,
				},
			},
		};
	}
	if (edit.kind === 'linked' && linkedId) {
		const partner = existingPartner;
		const basePartner = edit.basedOn && ownStyle(run.styles, edit.basedOn)?.linkedStyle;
		const { basedOn: _base, ...kept } = partner ?? {};
		const partnerBase =
			partner && edit.basedOn === old?.basedOn
				? partner.basedOn
				: basePartner && ownStyle(run.styles, basePartner)?.type === 'character'
					? basePartner
					: undefined;
		const formatting =
			partner && priorRun
				? applyRunChanges(partner.formatting, priorRun.formatting, edit.runFormatting)
				: edit.runFormatting;
		runStyles[linkedId] = {
			...kept,
			id: linkedId,
			type: 'character',
			name: partner?.name && partner.name !== old?.name + ' Char' ? partner.name : name + ' Char',
			linkedStyle: id,
			formatting,
			...(partner ? {} : { customStyle: true }),
			...(partnerBase && { basedOn: partnerBase }),
		};
	}
	return { catalogs: next, id, ...(edit.kind === 'linked' && linkedId && { linkedId }) };
}

function applyRunChanges(
	target: RunFormatting,
	before: RunFormatting,
	after: RunFormatting,
): RunFormatting {
	const result = { ...target };
	function patch<K extends keyof RunFormatting>(key: K) {
		const next = after[key];
		if (JSON.stringify(next) === JSON.stringify(before[key])) return;
		if (next === undefined) delete result[key];
		else result[key] = next;
	}
	for (const key of Object.keys({ ...before, ...after }) as Array<keyof RunFormatting>) patch(key);
	return result;
}
