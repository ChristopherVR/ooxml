import type { EditVsdxOptions, EditVsdxResult } from './edit';
import { openEditablePackage, writeEditedPackage } from './edit-package';
import { setCell } from './edit-geometry-cells';
import { serializeEditedXml } from './edit-text';
import { masterTemplate } from './edit-text-scope';
import { DEFAULTS, fail } from './package-common';
import { parseVsdx } from './parser';
import { visioXml } from './parts';
import { attribute, children } from './sheet';
import { visioLineJumpPaths } from './line-jumps';
import { executableCellFormula } from './cell-formula';
import { visioPageLayout } from './page-layout';
import type { VisioPage } from './model';

const jumpState = (page: VisioPage) =>
	JSON.stringify([
		Object.entries(visioPageLayout(page)).filter(([key]) => /^lineJump|^lineToLine/.test(key)),
		page.shapes.map((shape) => [
			shape.id,
			shape.lineJumpCode,
			shape.lineJumpStyle,
			shape.lineJumpDirectionX,
			shape.lineJumpDirectionY,
		]),
	]);

function plainRows(section: Element): boolean {
	return children(section, 'Row').every(
		(row) =>
			['MoveTo', 'LineTo', 'ArcTo'].includes(attribute(row, 'T') ?? '') &&
			Array.from(row.attributes).every((attr) => ['IX', 'T', 'Del'].includes(attr.name)) &&
			Array.from(row.children ?? row.childNodes)
				.filter((node) => node.nodeType === 1)
				.every((node) => {
					const cell = node as Element;
					return (
						cell.localName === 'Cell' &&
						cell.namespaceURI === section.namespaceURI &&
						['X', 'Y', ...(attribute(row, 'T') === 'ArcTo' ? ['A'] : [])].includes(
							attribute(cell, 'N') ?? '',
						) &&
						Array.from(cell.attributes).every((attr) => ['N', 'V', 'U', 'F'].includes(attr.name)) &&
						!executableCellFormula(attribute(cell, 'F')) &&
						!Array.from(cell.childNodes).some((child) => child.nodeType === 1)
					);
				}),
	);
}

/** Decode only the bounded M/L/semicircle grammar emitted by line-jumps. */
function cacheRows(path: string): { type: string; x: number; y: number; bow?: number }[] {
	const tokens = path.trim().split(/\s+/);
	const rows: ReturnType<typeof cacheRows> = [];
	while (tokens.length) {
		const command = tokens.shift();
		if (command === 'M' || command === 'L') {
			rows.push({
				type: command === 'M' ? 'MoveTo' : 'LineTo',
				x: Number(tokens.shift()),
				y: Number(tokens.shift()),
			});
		} else if (command === 'A') {
			const [rx, ry, rotation, large, sweep, x, y] = tokens.splice(0, 7).map(Number);
			if (
				rx === undefined ||
				ry !== rx ||
				rotation !== 0 ||
				large !== 0 ||
				(sweep !== 0 && sweep !== 1) ||
				x === undefined ||
				y === undefined
			)
				return [];
			rows.push({ type: 'ArcTo', x, y, bow: sweep ? rx : -rx });
		} else return [];
		if (rows.length > 10000) return [];
	}
	return rows.every((row) => Number.isFinite(row.x) && Number.isFinite(row.y)) ? rows : [];
}

/** Update explicit straight-route geometry so native Visio sees the edited decorations too. */
export async function refreshLineJumpCaches(
	pageIds: ReadonlySet<string>,
	source: Uint8Array | undefined,
	result: EditVsdxResult,
	options: EditVsdxOptions,
	deadline: number,
): Promise<EditVsdxResult> {
	if (!pageIds.size || !result.changedParts.length) return result;
	const check = () => {
		if (Date.now() >= deadline) fail('LIMIT_RUNTIME', 'Visio edit deadline exceeded.');
	};
	const limits = {
		...DEFAULTS,
		...options.limits,
		maxRuntimeMs: Math.max(1, deadline - Date.now()),
	};
	const { pkg, parts, pages } = await openEditablePackage(result.bytes, limits, check);
	const model = await parseVsdx(result.bytes, { limits });
	const before = source ? await parseVsdx(source, { limits }) : undefined;
	const template = masterTemplate(pkg);
	const changed = new Set(result.changedParts);
	const kept: string[] = [];
	// Geometry-row references can depend on row identity; never rewrite them silently.
	let referenced = false;
	for (const path of pkg.paths()) {
		check();
		if (!/^visio\/.*\.xml$/i.test(path)) continue;
		const root = await pkg.readXml(path);
		for (const node of Array.from(root.getElementsByTagName('*')))
			if (/\b(?:Geometry|Geom)\d*\./i.test(attribute(node, 'F') ?? '')) referenced = true;
	}
	for (const page of model.pages) {
		if (!pageIds.has(page.id)) continue;
		const prior = before?.pages.find((item) => item.id === page.id);
		if (prior && jumpState(prior) === jumpState(page)) continue;
		const path = pages.get(page.id)!;
		const root = await visioXml(pkg, path, 'PageContents');
		let dirty = false;
		for (const [shape, paths] of visioLineJumpPaths(page)) {
			check();
			if (shape.connectorRoute !== 'straight' || paths.length !== 1) continue;
			const target = children(children(root, 'Shapes')[0], 'Shape').find(
				(node) => attribute(node, 'ID') === shape.id,
			);
			if (!target) continue;
			const sections = children(target, 'Section').filter(
				(node) => attribute(node, 'N') === 'Geometry',
			);
			const section = sections[0];
			const rows = cacheRows(paths[0]!);
			if (
				referenced ||
				!section ||
				sections.length !== 1 ||
				!rows.length ||
				!plainRows(section) ||
				Array.from(section.getElementsByTagName('*')).some((node) =>
					executableCellFormula(attribute(node, 'F')),
				)
			) {
				kept.push(shape.id);
				continue;
			}
			const indices = new Map<string, string>();
			const masterId = attribute(target, 'Master');
			if (masterId) {
				const master = await template(masterId, attribute(target, 'MasterShape'));
				const base = children(master, 'Section').find(
					(node) =>
						attribute(node, 'N') === 'Geometry' &&
						attribute(node, 'IX') === attribute(section, 'IX'),
				);
				for (const row of children(base, 'Row'))
					indices.set(attribute(row, 'IX')!, attribute(row, 'T')!);
			}
			for (const row of children(section, 'Row')) section.removeChild(row);
			for (const [index, value] of rows.entries()) {
				const row = root.ownerDocument!.createElementNS(root.namespaceURI, 'Row');
				row.setAttribute('IX', String(index + 1));
				row.setAttribute('T', value.type);
				setCell(row, 'X', value.x);
				setCell(row, 'Y', value.y);
				if (value.bow !== undefined) setCell(row, 'A', value.bow);
				section.appendChild(row);
				indices.delete(String(index + 1));
			}
			for (const [index, type] of indices) {
				const row = root.ownerDocument!.createElementNS(root.namespaceURI, 'Row');
				row.setAttribute('IX', index);
				row.setAttribute('T', type);
				row.setAttribute('Del', '1');
				section.appendChild(row);
			}
			dirty = true;
		}
		if (dirty) {
			parts.set(path, serializeEditedXml(root, limits, check).bytes);
			changed.add(path);
		}
	}
	return {
		bytes: await writeEditedPackage(
			parts,
			options.maxOutputBytes ?? limits.maxInputBytes,
			deadline,
			check,
		),
		changedParts: [...changed].sort(),
		diagnostics: [
			...result.diagnostics,
			...(kept.length
				? [
						{
							code: 'edit-line-jump-cache-kept',
							message: `Connector ${kept.join(', ')} retains its saved geometry because formula dependencies prevent refreshing native jump caches.`,
						},
					]
				: []),
		],
	};
}
