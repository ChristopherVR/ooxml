import type { VisioGeometryEdit } from './edit-commands.js';
import type { VisioPackage } from './package.js';
import { analyzeVisioFormula } from './formula.js';
import { fail } from './package-common.js';
import { VISIO_NS, VISIO_LEGACY_NS, attribute, children } from './sheet.js';
import { createVisioDependencyQuery } from './edit-recalculate.js';
import { executableCellFormula, inertDoubleClickFormula } from './cell-formula.js';
import { assertVisioMasterIndependence } from './edit-master-scope.js';

const admitted = (node: Element) =>
	node.namespaceURI === VISIO_NS || node.namespaceURI === VISIO_LEGACY_NS;
function cells(sheet: Element): Map<string, Element> {
	const result = new Map<string, Element>();
	const add = (name: string, node: Element) => {
		const key = name.toLowerCase();
		if (result.has(key))
			fail('EDIT_UNSUPPORTED_PACKAGE_DEPENDENCY', 'Ambiguous inherited style cell.');
		result.set(key, node);
	};
	for (const node of children(sheet, 'Cell')) add(attribute(node, 'N') ?? '', node);
	for (const section of children(sheet, 'Section')) {
		const name = attribute(section, 'N') ?? '';
		const prefix =
			name === 'Geometry' ? `Geometry${Number(attribute(section, 'IX') ?? '0') + 1}` : name;
		for (const node of children(section, 'Cell'))
			add(`${prefix}.${attribute(node, 'N') ?? ''}`, node);
		for (const row of children(section, 'Row'))
			for (const node of children(row, 'Cell')) {
				const n = attribute(node, 'N') ?? '',
					ix = attribute(row, 'IX') ?? '0',
					rowName = attribute(row, 'N') ?? ix;
				add(
					name === 'Geometry'
						? `${prefix}.${n}${ix}`
						: n === 'Value' && ['User', 'Prop'].includes(name)
							? `${prefix}.${rowName}`
							: `${prefix}.${rowName}.${n}`,
					node,
				);
			}
	}
	return result;
}
/** Separate master and page ShapeSheet ID namespaces; refuse unsupported affected metadata and inheritance. */
export async function assertGeometryPackageScope(
	pkg: VisioPackage,
	pagePaths: ReadonlySet<string>,
	commands: readonly VisioGeometryEdit[],
	check: () => void,
	roots: ReadonlyMap<string, Element>,
): Promise<ReadonlySet<Element>> {
	const changed = commands.flatMap((command) =>
		command.type === 'delete-shape'
			? []
			: (command.type === 'move-shape'
					? ['PinX', 'PinY']
					: command.type === 'resize-shape'
						? ['Width', 'Height']
						: ['PinX', 'PinY', 'Width', 'Height']
				).map((cell) => ({ pageId: command.pageId, shapeId: command.shapeId, cell })),
	);
	const deleting = commands.filter((command) => command.type === 'delete-shape');
	let query: ReturnType<typeof createVisioDependencyQuery> | undefined;
	const hasMasterInstances = [...roots.values()].some((root) =>
		Array.from(root.getElementsByTagName('*')).some(
			(node) => admitted(node) && node.localName === 'Shape' && node.hasAttribute('Master'),
		),
	);
	const dependencyCache = new Map<string, boolean>();
	const depends = (pageId: string, shapeId: string, cell: string) => {
		const key = JSON.stringify([pageId, shapeId, cell.toLowerCase()]);
		let result = dependencyCache.get(key);
		if (result === undefined) {
			query ??= createVisioDependencyQuery(roots, { check });
			result = query({ pageId, shapeId, cell }, changed);
			dependencyCache.set(key, result);
		}
		return result;
	};
	const affected = (shapeId: string, cell: string, pageId?: string) => {
		const pages = pageId === undefined ? [...roots.keys()] : [pageId];
		return pages.some(
			(page) =>
				changed.some(
					(target) =>
						target.pageId === page &&
						target.shapeId === shapeId &&
						target.cell.toLowerCase() === cell.toLowerCase(),
				) ||
				deleting.some((target) => target.pageId === page && target.shapeId === shapeId) ||
				depends(page, shapeId, cell),
		);
	};
	const inspectFormula = (node: Element, path: string, pageId?: string, localShapeId?: string) => {
		const source = executableCellFormula(attribute(node, 'F'));
		if (!source) return;
		// These literal event handlers execute on double click, not geometry recalculation.
		if (inertDoubleClickFormula(attribute(node, 'N') ?? '', source)) return;
		let analysis: ReturnType<typeof analyzeVisioFormula>;
		try {
			analysis = analyzeVisioFormula(source);
		} catch (error) {
			fail(
				'EDIT_UNSUPPORTED_PACKAGE_DEPENDENCY',
				`Cannot analyze formula in ${path}: ${error instanceof Error ? error.message : 'unsupported formula'}`,
			);
		}
		if (
			analysis.dynamic ||
			analysis.references.some((ref) => {
				const shapeId = ref.shapeId ?? localShapeId;
				return shapeId !== undefined && affected(shapeId, ref.cell, pageId);
			})
		)
			fail(
				'EDIT_UNSUPPORTED_PACKAGE_DEPENDENCY',
				`Formula in ${path}, cell ${attribute(node, 'N') ?? node.localName} has affected or unknown dependencies outside editable page cells.`,
			);
	};
	let documentRoot: Element | undefined;
	for (const path of pkg.paths()) {
		if (!/^visio\/.*\.xml$/i.test(path) || pagePaths.has(path)) continue;
		// Active masters are checked below with their effective instance and template scopes.
		// Unused definitions remain untouched and cannot influence page geometry caches.
		if (/^visio\/masters\//i.test(path)) continue;
		const root = await pkg.readXml(path);
		const masterScope =
			/^visio\/masters\//i.test(path) ||
			(admitted(root) && ['MasterContents', 'Masters'].includes(root.localName));
		if (masterScope) continue;
		if (admitted(root) && root.localName === 'VisioDocument') {
			if (documentRoot)
				fail('EDIT_UNSUPPORTED_PACKAGE_DEPENDENCY', 'Multiple document scopes are ambiguous.');
			documentRoot = root;
		}
		const pending: { node: Element; pageId?: string; style: boolean }[] = [
			{ node: root, style: false },
		];
		while (pending.length) {
			check();
			const item = pending.pop()!,
				node = item.node;
			const pageId =
				admitted(node) && node.localName === 'Page' && root.localName === 'Pages'
					? attribute(node, 'ID')
					: item.pageId;
			const style = item.style || (admitted(node) && node.localName === 'StyleSheets');
			// Styles need instance context and local override analysis below.
			if (admitted(node) && !style) {
				inspectFormula(node, path, pageId);
				if (
					deleting.some(
						(target) =>
							(pageId === undefined || target.pageId === pageId) &&
							['FromSheet', 'ToSheet', 'ShapeID', 'SheetID'].some(
								(name) => attribute(node, name) === target.shapeId,
							),
					)
				)
					fail('EDIT_REFERENCED_DELETE', 'Scoped metadata references the deleted shape.');
			}
			for (const child of Array.from(node.childNodes))
				if (child.nodeType === 1)
					pending.push({
						node: child as Element,
						style,
						...(pageId === undefined ? {} : { pageId }),
					});
		}
	}
	let movePins: ReadonlySet<Element> = new Set();
	if (hasMasterInstances) {
		for (const command of commands) {
			const root = roots.get(command.pageId);
			if (!root) continue;
			const target = Array.from(root.getElementsByTagNameNS(root.namespaceURI!, 'Shape')).find(
				(shape) => attribute(shape, 'ID') === command.shapeId,
			);
			if (target?.hasAttribute('MasterShape') && !target.hasAttribute('Master'))
				fail('UNSUPPORTED_GEOMETRY_EDIT', 'Editing master-linked shapes is unsupported.');
			if (
				target?.hasAttribute('Master') &&
				(command.type !== 'move-shape' ||
					!['PinX', 'PinY', 'Width', 'Height', 'LocPinX', 'LocPinY'].every((name) =>
						children(target, 'Cell').some((node) => attribute(node, 'N') === name),
					))
			)
				fail('UNSUPPORTED_GEOMETRY_EDIT', 'Master moves require complete local transform caches.');
		}
		movePins = await assertVisioMasterIndependence(pkg, roots, commands, check);
	}
	if (!documentRoot) return movePins;
	const styles = new Map<string, Element>();
	for (const container of children(documentRoot, 'StyleSheets'))
		for (const style of children(container, 'StyleSheet')) {
			const id = attribute(style, 'ID');
			if (!id || styles.has(id))
				fail('EDIT_UNSUPPORTED_PACKAGE_DEPENDENCY', 'Style IDs must be unique.');
			styles.set(id, style);
		}
	const styleCache = new Map<string, Map<string, Element>>();
	const documentSheet = children(documentRoot, 'DocumentSheet')[0];
	const inherited = (
		id: string,
		category: string,
		seen = new Set<string>(),
	): Map<string, Element> => {
		check();
		const cacheKey = `${category}:${id}`;
		const cached = styleCache.get(cacheKey);
		if (cached) return cached;
		if (seen.has(id) || seen.size >= 64)
			fail('EDIT_UNSUPPORTED_PACKAGE_DEPENDENCY', 'Style ancestry is cyclic or too deep.');
		const style = styles.get(id);
		if (!style)
			fail('EDIT_UNSUPPORTED_PACKAGE_DEPENDENCY', 'Selected style ancestry cannot be resolved.');
		seen.add(id);
		const parent = attribute(style, category);
		const result =
			parent !== undefined && parent !== id
				? new Map(inherited(parent, category, seen))
				: new Map<string, Element>();
		for (const [name, node] of cells(style))
			if (attribute(node, 'F') !== 'Inh') result.set(name, node);
		styleCache.set(cacheKey, result);
		return result;
	};
	for (const [pageId, root] of roots) {
		const shapes = Array.from(root.getElementsByTagNameNS(root.namespaceURI!, 'Shape'));
		for (const command of commands)
			if (
				command.type === 'create-rectangle' &&
				command.pageId === pageId &&
				!shapes.some((shape) => attribute(shape, 'ID') === command.shapeId)
			) {
				// Detached admission context only: a planned rectangle receives document default styles.
				const shape = root.ownerDocument!.createElementNS(root.namespaceURI, 'Shape');
				shape.setAttribute('ID', command.shapeId);
				for (const name of ['PinX', 'PinY', 'Width', 'Height', 'LocPinX', 'LocPinY', 'Angle']) {
					const node = root.ownerDocument!.createElementNS(root.namespaceURI, 'Cell');
					node.setAttribute('N', name);
					shape.appendChild(node);
				}
				shapes.push(shape);
			}
		for (const shape of shapes) {
			check();
			const shapeId = attribute(shape, 'ID');
			if (!shapeId) continue;
			// Effective formulas owned solely by a removed instance disappear with that instance.
			// Other instances and package metadata still undergo reference checks.
			if (deleting.some((command) => command.pageId === pageId && command.shapeId === shapeId))
				continue;
			const local = cells(shape);
			for (const category of ['FillStyle', 'LineStyle', 'TextStyle']) {
				const styleId =
					attribute(shape, category) ??
					attribute(documentSheet, category) ??
					(styles.has('0') ? '0' : undefined);
				if (styleId === undefined) continue;
				for (const [name, node] of inherited(styleId, category)) {
					const override = local.get(name);
					if (override && attribute(override, 'F') !== 'Inh') continue;
					// The page graph supplies defaults only for truly absent effective cells.
					// An inherited cache must not be replaced by a synthetic dimension formula.
					if (depends(pageId, shapeId, name))
						fail(
							'EDIT_UNSUPPORTED_PACKAGE_DEPENDENCY',
							'Affected inherited style cells cannot use synthetic page defaults.',
						);
					inspectFormula(node, `inherited ${category} ${styleId}`, pageId, shapeId);
				}
			}
		}
	}
	return movePins;
}
