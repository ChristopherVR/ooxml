import { parseInteger, parseOnOff } from './simple-types';
import { children, first, getW, makeW, WORD_NS, type XmlElement } from './xml';

interface LatentStyleFlags {
	locked?: boolean;
	semiHidden?: boolean;
	unhideWhenUsed?: boolean;
	qFormat?: boolean;
}

export interface LatentStyleException extends LatentStyleFlags {
	name: string;
	uiPriority?: number;
}

export interface LatentStyles {
	count?: number;
	defaults: LatentStyleFlags & { uiPriority?: number };
	exceptions: LatentStyleException[];
}

const FLAGS = [
	['locked', 'defLockedState'],
	['semiHidden', 'defSemiHidden'],
	['unhideWhenUsed', 'defUnhideWhenUsed'],
	['qFormat', 'defQFormat'],
] as const;

function readFlags(element: XmlElement, defaults: boolean): LatentStyles['defaults'] {
	const flags: LatentStyles['defaults'] = {};
	for (const [key, defaultAttr] of FLAGS) {
		const value = parseOnOff(getW(element, defaults ? defaultAttr : key));
		if (value !== undefined) flags[key] = value;
	}
	const uiPriority = parseInteger(getW(element, defaults ? 'defUIPriority' : 'uiPriority'));
	if (uiPriority !== undefined) flags.uiPriority = uiPriority;
	return flags;
}

export function readLatentStyles(root: XmlElement): LatentStyles | undefined {
	const latent = first(root, 'latentStyles');
	if (!latent) return undefined;
	const count = parseInteger(getW(latent, 'count'));
	return {
		...(count !== undefined && { count }),
		defaults: readFlags(latent, true),
		exceptions: children(latent, 'lsdException').flatMap((element) => {
			const name = getW(element, 'name');
			return name === undefined ? [] : [{ name, ...readFlags(element, false) }];
		}),
	};
}

function patchAttr(
	element: XmlElement,
	key: string,
	value: number | boolean | undefined,
	prior: typeof value,
) {
	if (value === prior) return;
	if (value === undefined) element.removeAttributeNS(WORD_NS, key);
	else
		element.setAttributeNS(
			WORD_NS,
			`w:${key}`,
			String(typeof value === 'boolean' ? Number(value) : value),
		);
}

function writeFlags(
	element: XmlElement,
	flags: LatentStyles['defaults'],
	prior: LatentStyles['defaults'] | undefined,
	defaults: boolean,
) {
	for (const [key, defaultAttr] of FLAGS)
		patchAttr(element, defaults ? defaultAttr : key, flags[key], prior?.[key]);
	patchAttr(
		element,
		defaults ? 'defUIPriority' : 'uiPriority',
		flags.uiPriority,
		prior?.uiPriority,
	);
}

export function writeLatentStyles(
	root: XmlElement,
	next: LatentStyles | undefined,
	prior: LatentStyles | undefined,
): void {
	if (JSON.stringify(next) === JSON.stringify(prior)) return;
	let latent = first(root, 'latentStyles');
	if (!next) {
		if (latent) root.removeChild(latent);
		return;
	}
	const doc = root.ownerDocument;
	if (!latent) {
		latent = makeW(doc, 'latentStyles');
		root.insertBefore(latent, first(root, 'style') ?? null);
	}
	patchAttr(latent, 'count', next.count, prior?.count);
	writeFlags(latent, next.defaults, prior?.defaults, true);
	const original = children(latent, 'lsdException').filter(
		(element) => getW(element, 'name') !== undefined,
	);
	const byName = new Map<string, number[]>();
	for (const [index, entry] of original.entries()) {
		const name = getW(entry, 'name') ?? '';
		const indexes = byName.get(name) ?? [];
		indexes.push(index);
		byName.set(name, indexes);
	}
	for (const indexes of byName.values()) indexes.reverse();
	const used = new Set<XmlElement>();
	const ordered = next.exceptions.map((entry) => {
		const index = byName.get(entry.name)?.pop();
		const element =
			(index === undefined ? undefined : original[index]) ?? makeW(doc, 'lsdException');
		used.add(element);
		if (index === undefined) element.setAttributeNS(WORD_NS, 'w:name', entry.name);
		writeFlags(element, entry, index === undefined ? undefined : prior?.exceptions[index], false);
		return element;
	});
	if (
		ordered.length === original.length &&
		ordered.every((element, index) => element === original[index])
	)
		return;
	const anchor = original.at(-1)?.nextSibling ?? null;
	for (const element of ordered) latent.insertBefore(element, anchor);
	for (const element of original) if (!used.has(element)) latent.removeChild(element);
}
