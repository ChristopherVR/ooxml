import { attribute, child, children } from './sheet';
import { indexedPart, related, visioXml } from './parts';
import type { VisioPackage } from './package';
import { fail } from './package-common';
import { parseVisioFormula, type VisioFormulaAst } from './formula';
import { indexCells, type VisioCellKey } from './edit-recalculate-index';
import { recalculateVisioCells } from './edit-recalculate';

function pageFunctions(source: string): Set<string> {
	const result = new Set<string>();
	if (!/\bPAGE(?:NUMBER|COUNT)\s*\(/i.test(source)) return result;
	const pending: VisioFormulaAst[] = [parseVisioFormula(source)];
	while (pending.length) {
		const node = pending.pop()!;
		if (node.kind === 'call') {
			if (node.name === 'PAGENUMBER' || node.name === 'PAGECOUNT') result.add(node.name);
			pending.push(...node.args);
		} else if (node.kind === 'binary') pending.push(node.left, node.right);
		else if (node.kind === 'unary') pending.push(node.operand);
	}
	return result;
}
const background = (page: Element) => ['1', 'true'].includes(attribute(page, 'Background') ?? '');
function contexts(pages: Element) {
	const list = children(pages, 'Page');
	const pageCount = list.filter((page) => !background(page)).length;
	let pageNumber = 0;
	return new Map(
		list.map((page) => [
			attribute(page, 'ID')!,
			{
				pageNumber: background(page) ? 0 : ++pageNumber,
				pageCount,
			},
		]),
	);
}

/** Refresh local numeric page functions and their proven static dependency closure atomically. */
export async function recalculatePageFormulas(
	pkg: VisioPackage,
	pagesPart: string,
	pages: Element,
	paths: ReadonlyMap<string, string>,
	dirty: Map<string, Element>,
	check: () => void,
): Promise<void> {
	const before = contexts(await visioXml(pkg, pagesPart, 'Pages'));
	const after = contexts(pages);
	const roots = new Map<string, Element>();
	const sheets = new Map<string, Element>();
	let needed = false;
	for (const page of children(pages, 'Page')) {
		check();
		const id = attribute(page, 'ID')!,
			path = paths.get(id)!;
		const source = dirty.get(path) ?? (await visioXml(pkg, path, 'PageContents'));
		const root = (source.ownerDocument!.cloneNode(true) as Document).documentElement;
		const sheet = child(page, 'PageSheet');
		if (sheet) {
			const clone = sheet.cloneNode(true) as Element;
			root.appendChild(clone);
			sheets.set(id, clone);
		}
		roots.set(id, root);
		for (const node of Array.from(root.getElementsByTagName('*'))) {
			check();
			const functions = pageFunctions(attribute(node, 'F') ?? '');
			const old = before.get(id),
				current = after.get(id)!;
			if (
				(functions.has('PAGENUMBER') && old?.pageNumber !== current.pageNumber) ||
				(functions.has('PAGECOUNT') && old?.pageCount !== current.pageCount)
			)
				needed = true;
		}
	}
	// Context-dependent master/style formulas cannot be materialized as local caches yet.
	const documentPart = (await related(pkg, '', 'document'))!;
	const document = await visioXml(pkg, documentPart, 'VisioDocument');
	const inherited = [document];
	const mastersPart = await related(pkg, documentPart, 'masters', false);
	if (mastersPart) {
		for (const master of children(await visioXml(pkg, mastersPart, 'Masters'), 'Master')) {
			check();
			inherited.push(
				await visioXml(
					pkg,
					await indexedPart(pkg, mastersPart, master, 'master'),
					'MasterContents',
				),
			);
		}
	}
	for (const root of inherited)
		for (const node of Array.from(root.getElementsByTagName('*'))) {
			check();
			if (pageFunctions(attribute(node, 'F') ?? '').size)
				fail(
					'EDIT_UNSUPPORTED_PAGE_FORMULA',
					'Inherited page-dependent formula caches cannot be safely refreshed.',
				);
		}
	if (!needed) return;
	const options = { check, pageContext: after };
	const changed: VisioCellKey[] = [];
	for (const cell of indexCells(roots, options).values()) {
		if (!cell.node) continue;
		const functions = pageFunctions(attribute(cell.node, 'F') ?? '');
		const old = before.get(cell.pageId),
			current = after.get(cell.pageId)!;
		if (
			(functions.has('PAGENUMBER') && old?.pageNumber !== current.pageNumber) ||
			(functions.has('PAGECOUNT') && old?.pageCount !== current.pageCount)
		)
			changed.push(cell);
	}
	for (const id of recalculateVisioCells(roots, changed, options)) {
		const root = roots.get(id)!;
		const sheet = sheets.get(id);
		if (sheet) {
			root.removeChild(sheet);
			const page = children(pages, 'Page').find((node) => attribute(node, 'ID') === id)!;
			const original = child(page, 'PageSheet')!;
			page.replaceChild(sheet, original);
			dirty.set(pagesPart, pages);
		}
		dirty.set(paths.get(id)!, root);
	}
}
