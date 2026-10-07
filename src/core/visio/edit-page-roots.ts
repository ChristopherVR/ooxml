import type { VisioPackage } from './package';
import { indexedPart, related, visioXml } from './parts';
import { children } from './sheet';

/** Shared inventory of formula-bearing parts; OPC paths come from relationships. */
export async function editableVisioPageRoots(
	pkg: VisioPackage,
	pagesPart: string,
	pages: Element,
	pagePaths: ReadonlyMap<string, string>,
	dirty: ReadonlyMap<string, Element>,
	check: () => void,
	excludePageId?: string,
): Promise<Map<string, Element>> {
	const admitted = new Map(
		[...pagePaths].filter(([id]) => id !== excludePageId).map(([, path]) => [path, 'PageContents']),
	);
	admitted.set(pagesPart, 'Pages');
	const documentPart = (await related(pkg, '', 'document'))!;
	admitted.set(documentPart, 'VisioDocument');
	const mastersPart = await related(pkg, documentPart, 'masters', false);
	if (mastersPart) {
		admitted.set(mastersPart, 'Masters');
		for (const master of children(await visioXml(pkg, mastersPart, 'Masters'), 'Master')) {
			check();
			admitted.set(await indexedPart(pkg, mastersPart, master, 'master'), 'MasterContents');
		}
	}
	const roots = new Map<string, Element>();
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
	return roots;
}
