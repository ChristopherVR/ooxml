import { attribute, children } from './sheet';
import { fail } from './package-common';

/** Native Visio 2013+ FaceNames use names; older source fixtures can carry numeric IDs. */
export function formattingFont(
	document: Element,
	family: string,
): { value: string; formula?: string } {
	const containers = children(document, 'FaceNames');
	if (containers.length !== 1)
		fail('UNSUPPORTED_FORMAT_EDIT', 'Font family requires existing document FaceNames.');
	const fonts = children(containers[0], 'FaceName');
	const ids = new Set<string>();
	for (const font of fonts) {
		const id = attribute(font, 'ID');
		if (id === undefined) continue;
		if (!/^\d+$/.test(id) || ids.has(id))
			fail('UNSUPPORTED_FORMAT_EDIT', 'Explicit document font IDs must be numeric and unique.');
		ids.add(id);
	}
	const matches = fonts.filter(
		(font) =>
			(attribute(font, 'Name') ?? attribute(font, 'NameU'))?.toLowerCase() === family.toLowerCase(),
	);
	if (matches.length !== 1)
		fail('UNSUPPORTED_FORMAT_EDIT', 'Font family must match one existing document FaceName.');
	const font = matches[0]!,
		id = attribute(font, 'ID');
	if (id !== undefined) return { value: id };
	const name = attribute(font, 'Name') ?? attribute(font, 'NameU')!;
	return { value: name, formula: `FONT("${name.replaceAll('"', '""')}")` };
}
