// Ordered picture effects. Opaque effects are preserved and explicitly reported to painters.
import { NS, buildXml, elements, parseXml, type XmlElement } from '../xml/index';

export type DrawingPictureEffect = (
	| { kind: 'grayscale' }
	| { kind: 'luminance'; brightness?: number; contrast?: number }
	| { kind: 'alphaModFix'; amount: number }
	| { kind: 'unsupported'; name: string }
) & { sourceXml?: string };

/** A DrawingML percentage as a fraction, including the Strict percent spelling. */
export function picturePercentage(value: unknown): number | undefined {
	const text = String(value ?? '').trim();
	if (!text) return undefined;
	const percent = text.endsWith('%');
	const n = Number(percent ? text.slice(0, -1) : text) / (percent ? 100 : 100000);
	return Number.isFinite(n) ? n : undefined;
}

/** Shared semantic adapter for DOM and PowerPoint source-tree attributes. */
export function readPictureEffect(
	name: string,
	attribute: (name: string) => unknown,
): DrawingPictureEffect | undefined {
	if (name === 'grayscl') return { kind: 'grayscale' };
	if (name === 'lum') {
		const brightness = picturePercentage(attribute('bright'));
		const contrast = picturePercentage(attribute('contrast'));
		return {
			kind: 'luminance',
			...(brightness !== undefined ? { brightness } : {}),
			...(contrast !== undefined ? { contrast } : {}),
		};
	}
	if (name === 'alphaModFix') {
		const raw = attribute('amt');
		const amount = raw === undefined || raw === null ? 1 : picturePercentage(raw);
		if (amount !== undefined && amount >= 0) return { kind: 'alphaModFix', amount };
	}
	return undefined;
}

export function parsePictureEffects(blip: XmlElement | undefined): DrawingPictureEffect[] {
	if (!blip) return [];
	return elements(blip).map((node) => ({
		...((node.namespaceURI === NS.a
			? readPictureEffect(node.localName, (name) => node.getAttribute(name))
			: undefined) ?? { kind: 'unsupported' as const, name: node.localName }),
		sourceXml: buildXml(node),
	}));
}

/** Replaces the effect sequence in authored order, retaining unknown XML and supported-node metadata. */
export function applyPictureEffects(
	blip: XmlElement,
	effects: readonly DrawingPictureEffect[],
): void {
	if (JSON.stringify(parsePictureEffects(blip)) === JSON.stringify(effects)) return;
	const nodes = effects.map((effect) => {
		const name =
			effect.kind === 'grayscale'
				? 'grayscl'
				: effect.kind === 'luminance'
					? 'lum'
					: effect.kind === 'unsupported'
						? effect.name
						: effect.kind;
		const source = effect.sourceXml ? parseXml(effect.sourceXml).documentElement : undefined;
		if (effect.kind === 'unsupported' && !source)
			throw new Error('An unsupported picture effect needs source XML');
		const node = source
			? (blip.ownerDocument.importNode(source, true) as XmlElement)
			: blip.ownerDocument.createElementNS(NS.a, `a:${name}`);
		if (effect.kind !== 'unsupported' && (node.namespaceURI !== NS.a || node.localName !== name))
			throw new Error('Picture effect source does not match its kind');
		const set = (name: string, value: number | undefined) => {
			const original = picturePercentage(node.getAttribute(name));
			if (value === undefined) node.removeAttribute(name);
			else if (value !== original) {
				if (!Number.isFinite(value)) throw new RangeError('Picture effect must be finite');
				node.setAttribute(name, String(Math.round(value * 100000)));
			}
		};
		if (effect.kind === 'luminance') {
			set('bright', effect.brightness);
			set('contrast', effect.contrast);
		}
		if (effect.kind === 'alphaModFix') {
			if (effect.amount < 0) throw new RangeError('Picture alpha must be nonnegative');
			set('amt', effect.amount);
		}
		return node;
	});
	for (const node of elements(blip)) blip.removeChild(node);
	for (const node of nodes) blip.appendChild(node);
}

/** RGB linear-transfer approximation used by the shared and PowerPoint SVG painters. */
export function pictureLuminanceTransfer(
	brightness = 0,
	contrast = 0,
): { slope: number; intercept: number } {
	const round = (value: number) => Math.round(value * 1e6) / 1e6;
	const c = 1 + contrast;
	return { slope: round(c), intercept: round(brightness + ((1 - brightness) * (1 - c)) / 2) };
}

export interface PictureEffectPatch {
	grayscale?: boolean;
	brightness?: number;
	contrast?: number;
	alpha?: number;
}
/** Edits supported fields while retaining opaque operations and refusing ambiguous duplicate targets. */
export function patchPictureEffects(
	source: readonly DrawingPictureEffect[],
	patch: PictureEffectPatch,
): DrawingPictureEffect[] {
	const result = structuredClone(source) as DrawingPictureEffect[];
	const upsert = (
		kind: DrawingPictureEffect['kind'],
		update: (effect: DrawingPictureEffect | undefined) => DrawingPictureEffect | undefined,
	) => {
		const matches = result.flatMap((effect, index) => (effect.kind === kind ? [index] : []));
		if (matches.length > 1) throw new Error(`Ambiguous duplicate picture effect: ${kind}`);
		const index = matches[0];
		const value = update(index === undefined ? undefined : result[index]);
		if (index !== undefined) {
			if (value) result[index] = value;
			else result.splice(index, 1);
		} else if (value) {
			const ext = result.findIndex(
				(effect) => effect.kind === 'unsupported' && effect.name === 'extLst',
			);
			result.splice(ext < 0 ? result.length : ext, 0, value);
		}
	};
	for (const key of ['brightness', 'contrast', 'alpha'] as const)
		if (patch[key] !== undefined && !Number.isFinite(patch[key]))
			throw new RangeError('Picture effect must be finite');
	if (patch.alpha !== undefined && patch.alpha < 0)
		throw new RangeError('Picture alpha must be nonnegative');
	if (patch.grayscale !== undefined)
		upsert('grayscale', (effect) =>
			patch.grayscale ? (effect ?? { kind: 'grayscale' }) : undefined,
		);
	if (patch.brightness !== undefined || patch.contrast !== undefined)
		upsert('luminance', (effect) => ({
			...(effect?.kind === 'luminance' ? effect : { kind: 'luminance' as const }),
			...(patch.brightness !== undefined ? { brightness: patch.brightness } : {}),
			...(patch.contrast !== undefined ? { contrast: patch.contrast } : {}),
		}));
	if (patch.alpha !== undefined)
		upsert('alphaModFix', (effect) => ({
			...(effect?.kind === 'alphaModFix' ? effect : { kind: 'alphaModFix' as const }),
			amount: patch.alpha!,
		}));
	return result;
}
