import { attribute, children } from './sheet.js';
import { fail } from './package-common.js';
import type { VisioPackage } from './package.js';
import { indexedPart, related, visioXml } from './parts.js';

const escape = (value: string) => value.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
/** Transform formula syntax while leaving quoted string literals byte-for-byte intact. */
function syntax(source: string, transform: (segment: string) => string): string {
	return source
		.split(/("(?:[^"]|"")*")/g)
		.map((part, index) => (index % 2 ? part : transform(part)))
		.join('');
}
const unquoted = (source: string) =>
	source
		.split(/("(?:[^"]|"")*")/g)
		.map((part, index) => (index % 2 ? ' ' : part))
		.join('');
function localCell(node: Element): string | undefined {
	const row = node.parentNode as Element | null;
	const section = row?.parentNode as Element | null;
	if (row?.localName === 'Row' && section?.localName === 'Section') {
		const name = attribute(section, 'N'),
			rowName = attribute(row, 'N') ?? attribute(row, 'IX');
		if (!name || !rowName) return undefined;
		return `${name}.${rowName}${attribute(node, 'N') === 'Value' && ['User', 'Prop'].includes(name) ? '' : `.${attribute(node, 'N')}`}`;
	}
	return attribute(node, 'N');
}

/** Native local/universal rename semantics, static page references and direct page-name caches. */
export async function renameVisioPage(
	pkg: VisioPackage,
	pagesPart: string,
	pages: Element,
	pagePaths: ReadonlyMap<string, string>,
	dirty: Map<string, Element>,
	pageId: string,
	name: string,
	check: () => void,
): Promise<void> {
	const page = children(pages, 'Page').find((node) => attribute(node, 'ID') === pageId);
	if (!page) fail('EDIT_TARGET_NOT_FOUND', 'Page does not exist.');
	const oldName = attribute(page, 'Name') ?? attribute(page, 'NameU') ?? '';
	if (oldName === name) return;
	const oldUniversal = attribute(page, 'NameU') ?? oldName;
	const customUniversal = ['1', 'true'].includes(attribute(page, 'IsCustomNameU') ?? '');
	const universal = customUniversal ? oldUniversal : name;
	for (const other of children(pages, 'Page')) {
		if (other === page) continue;
		if (
			['Name', 'NameU'].some((key) => {
				const value = attribute(other, key)?.toLowerCase();
				return value === name.toLowerCase() || value === universal.toLowerCase();
			})
		)
			fail('EDIT_DUPLICATE_PAGE_NAME', 'Page name already exists.');
	}
	const roots = new Map<string, Element>();
	const admitted = new Map([...pagePaths.values()].map((path) => [path, 'PageContents']));
	admitted.set(pagesPart, 'Pages');
	const documentPart = (await related(pkg, '', 'document'))!;
	admitted.set(documentPart, 'VisioDocument');
	const mastersPart = await related(pkg, documentPart, 'masters', false);
	if (mastersPart) {
		admitted.set(mastersPart, 'Masters');
		for (const master of children(await visioXml(pkg, mastersPart, 'Masters'), 'Master'))
			admitted.set(await indexedPart(pkg, mastersPart, master, 'master'), 'MasterContents');
	}
	for (const [path, expected] of admitted) {
		check();
		const source =
			path === pagesPart ? pages : (dirty.get(path) ?? (await visioXml(pkg, path, expected)));
		roots.set(
			path,
			path === pagesPart
				? pages
				: (source.ownerDocument!.cloneNode(true) as Document).documentElement,
		);
	}
	const changedCells: { node: Element; cell: string }[] = [];
	const prefix = new RegExp(`Pages\\[${escape(oldUniversal)}\\]!`, 'gi');
	const targetByPath = new Map([...pagePaths].map(([id, path]) => [path, id]));
	for (const [path, root] of roots) {
		let changed = false;
		for (const node of Array.from(root.getElementsByTagName('*'))) {
			check();
			const formula = attribute(node, 'F');
			if (formula) {
				const code = unquoted(formula);
				let target = targetByPath.get(path);
				if (path === pagesPart) {
					let ancestor: Node | null = node;
					while (ancestor && (ancestor as Element).localName !== 'Page')
						ancestor = ancestor.parentNode;
					target = ancestor ? attribute(ancestor as Element, 'ID') : undefined;
				}
				const direct = /^=?\s*(?:Pages\[([^\]]+)\]!ThePage!)?PAGENAME\(\s*(0|750)?\s*\)\s*$/i.exec(
					code,
				);
				const isTarget =
					direct?.[1] !== undefined
						? direct[1].toLowerCase() === oldUniversal.toLowerCase()
						: target === pageId;
				if (direct && isTarget) {
					if (node.localName !== 'Cell' || node.hasAttribute('E'))
						fail(
							'EDIT_UNSUPPORTED_PAGE_NAME_FORMULA',
							'Page-name cache has an unsupported scope or error.',
						);
					const value = direct[2] === '750' ? universal : name;
					if (attribute(node, 'V') !== value) {
						node.setAttribute('V', value);
						changed = true;
						const cell = localCell(node);
						if (cell) changedCells.push({ node, cell });
					}
				} else if (/\bPAGENAME\s*\(/i.test(code)) {
					if (
						target === pageId ||
						!target ||
						new RegExp(`Pages\\[${escape(oldUniversal)}\\]!`, 'i').test(code)
					)
						fail(
							'EDIT_UNSUPPORTED_PAGE_NAME_FORMULA',
							'Complex page-name formula caches cannot be safely refreshed.',
						);
				}
				if (universal !== oldUniversal) {
					const updated = syntax(formula, (part) =>
						part.replace(prefix, () => `Pages[${universal}]!`),
					);
					if (updated !== formula) {
						if (/[\]\r\n]/.test(universal))
							fail(
								'EDIT_UNSUPPORTED_PAGE_REFERENCE',
								'Page name cannot be encoded in a static page reference.',
							);
						node.setAttribute('F', updated);
						changed = true;
					}
				}
			}
			// Literal hyperlink destinations use the local page name, optionally followed by a shape.
			if (node.localName === 'Cell' && attribute(node, 'N') === 'SubAddress') {
				const row = node.parentNode as Element,
					section = row.parentNode as Element;
				const value = attribute(node, 'V') ?? '';
				if (
					attribute(section, 'N') === 'Hyperlink' &&
					(!formula || /^"(?:[^"]|"")*"$/.test(formula)) &&
					(value.toLowerCase() === oldName.toLowerCase() ||
						value.toLowerCase().startsWith(`${oldName.toLowerCase()}/`))
				) {
					const updated = name + value.slice(oldName.length);
					node.setAttribute('V', updated);
					if (formula) node.setAttribute('F', `"${updated.replaceAll('"', '""')}"`);
					changed = true;
				}
			}
		}
		if (changed) dirty.set(path, root);
	}
	// Numeric recalculation cannot safely refresh string-valued dependent formulas yet.
	for (const root of roots.values())
		for (const node of Array.from(root.getElementsByTagName('*'))) {
			check();
			const formula = attribute(node, 'F');
			if (!formula) continue;
			if (
				changedCells.length &&
				/\b(?:INDIRECT|EVALCELL|EVALTEXT|GETREF|REF|SETATREF|SETATREFEXPR|SETATREFEVAL|CALL|RUNADDON)\s*\(/i.test(
					unquoted(formula),
				)
			)
				fail(
					'EDIT_UNSUPPORTED_PAGE_NAME_DEPENDENCY',
					'Dynamic page-name dependencies cannot be safely refreshed.',
				);
			for (const changed of changedCells) {
				if (changed.node === node) continue;
				const reference = new RegExp(
					`(?:^|[^a-z0-9_.])${escape(changed.cell)}(?:$|[^a-z0-9_.])`,
					'i',
				);
				if (reference.test(unquoted(formula)))
					fail(
						'EDIT_UNSUPPORTED_PAGE_NAME_DEPENDENCY',
						'String-valued page-name dependencies cannot be safely refreshed.',
					);
			}
		}
	page.setAttribute('Name', name);
	page.setAttribute('IsCustomName', '1');
	if (!customUniversal) {
		page.setAttribute('NameU', universal);
		page.setAttribute('IsCustomNameU', '1');
	}
	dirty.set(pagesPart, pages);
}
