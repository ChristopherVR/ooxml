// Splits the content part of a page or a master (`PageContents`, `MasterContents`) into a shell
// (the part with markers where the shape and connect lists were), the keyed shapes and the connects.
import type { DrawingKeys } from './keys';
import { splitShape } from './pieces';
import { attrOf, scanContainer } from './scan';
import { CONNECTS_MARK, SHAPES_MARK } from './schema';
import type { ConnectSnapshot, ShapeSnapshot } from './snapshot';
import type { ConnectEntry } from './schema';

/** Replaces the inner content of the first `name` element by `mark`; returns its children. */
function splitList(text: string, name: string, mark: string, child: string) {
	const box = scanContainer(text, name);
	if (!box) return { text, items: [] as string[] };
	const items = box.children.filter((c) => c.name === child).map((c) => text.slice(c.start, c.end));
	if (box.selfClosing) {
		const tag = text.slice(box.start, box.end);
		const qualified = /^<([^\s/>]+)/.exec(tag)?.[1] ?? name;
		const open = tag.replace(/\s*\/>$/, '>');
		return {
			text: text.slice(0, box.start) + open + mark + `</${qualified}>` + text.slice(box.end),
			items,
		};
	}
	return { text: text.slice(0, box.openEnd) + mark + text.slice(box.closeStart), items };
}

/** Pairs of id and shape key for the `Sheet.<id>` references in the F attribute of a cell. */
function formulaRefs(text: string, owner: ReadonlyMap<string, string>): string[] {
	const out: string[] = [];
	const seen = new Set<string>();
	for (const f of text.matchAll(/\sF\s*=\s*("[^"]*"|'[^']*')/g))
		for (const m of f[1]!.matchAll(/Sheet\.(\d+)(?!\d)/g)) {
			const id = m[1]!;
			const key = owner.get(id);
			if (key !== undefined && !seen.has(id)) {
				seen.add(id);
				out.push(id, key);
			}
		}
	return out;
}

function flatten(
	xml: string,
	parent: string | undefined,
	page: string,
	keys: DrawingKeys,
	out: ShapeSnapshot[],
): void {
	const split = splitShape(xml);
	const id = split ? attrOf(split.open, 'ID') : undefined;
	if (!split || id === undefined) return;
	const key = keys.shapeKey(page, id);
	out.push({
		key,
		id,
		...(parent !== undefined ? { parent } : {}),
		open: split.open,
		pieces: split.pieces,
	});
	for (const nested of split.nested) flatten(nested, key, page, keys, out);
}

/** The shell, shapes and connects of one content part, with keys minted for `scope`. */
export function splitContents(
	content: string,
	scope: string,
	keys: DrawingKeys,
): { shell: string; shapes: ShapeSnapshot[]; connects: ConnectSnapshot[] } {
	const shapesSplit = splitList(content, 'Shapes', SHAPES_MARK, 'Shape');
	const connectsSplit = splitList(shapesSplit.text, 'Connects', CONNECTS_MARK, 'Connect');
	const shapes: ShapeSnapshot[] = [];
	for (const text of shapesSplit.items) flatten(text, undefined, scope, keys, shapes);
	const owner = new Map(shapes.map((s) => [s.id, s.key]));
	for (const shape of shapes)
		for (const piece of shape.pieces)
			if (piece.kind === 'leaf') {
				const refs = formulaRefs(piece.text, owner);
				if (refs.length) piece.refs = refs;
			}
	const connects: ConnectSnapshot[] = [];
	for (const text of connectsSplit.items) {
		const a = (name: string) => attrOf(text, name) ?? '';
		const entry: ConnectEntry = {
			fk: owner.get(a('FromSheet')) ?? '',
			fi: a('FromSheet'),
			tk: owner.get(a('ToSheet')) ?? '',
			ti: a('ToSheet'),
			fc: a('FromCell'),
			tc: a('ToCell'),
			fp: a('FromPart'),
			tp: a('ToPart'),
		};
		connects.push({
			key: `${scope}/${entry.fk}.${entry.fi}.${entry.fc}.${entry.fp}>${entry.tk}.${entry.ti}.${entry.tc}.${entry.tp}`,
			entry,
		});
	}
	return { shell: connectsSplit.text, shapes, connects };
}
