import type { PptxCompatibilityWarning, XmlObject } from '../types';
import { isAlternateContentChoiceSupported } from '../utils/alternate-content';

interface AlternateContentWarning {
	code: string;
	message: string;
	severity: PptxCompatibilityWarning['severity'];
	scope: 'presentation' | 'slide';
	slideId?: string | undefined;
	xmlPath: string;
}

export function inspectAlternateContentWarnings(
	node: unknown,
	scope: 'presentation' | 'slide',
	slideId: string | undefined,
	path: string,
	report: (warning: AlternateContentWarning) => void,
): void {
	if (!node || typeof node !== 'object') {
		return;
	}
	if (Array.isArray(node)) {
		node.forEach((item, index) =>
			inspectAlternateContentWarnings(item, scope, slideId, `${path}[${index}]`, report),
		);
		return;
	}
	for (const [key, value] of Object.entries(node as XmlObject)) {
		if (key === 'mc:AlternateContent') {
			for (const block of asArray(value)) {
				for (const choice of asArray(block?.['mc:Choice'])) {
					if (isAlternateContentChoiceSupported(choice)) {
						continue;
					}
					if (isLegacyChoiceWithFallback(choice, block)) {
						continue;
					}
					const requires = String(choice['@_Requires'] || '(missing)');
					report({
						code: 'UNSUPPORTED_ALTERNATE_CONTENT_CHOICE',
						message: `An mc:Choice requiring "${requires}" is not implemented; its fallback is used when available.`,
						severity: block?.['mc:Fallback'] ? 'info' : 'warning',
						scope,
						slideId,
						xmlPath: `${path}/mc:AlternateContent/mc:Choice`,
					});
				}
			}
		}
		inspectAlternateContentWarnings(value, scope, slideId, `${path}/${key}`, report);
	}
}

/**
 * `Requires` tokens that name an OLDER format than the `mc:Fallback`
 * beside them, so taking the Fallback loses nothing.
 *
 * PowerPoint writes every OLE object and ActiveX control as an `mc:Choice
 * Requires="v"` (VML, the pre-2007 drawing format) next to an `mc:Fallback`
 * holding the same `p:oleObj` / `p:control` plus its DrawingML preview
 * `p:pic`. The reader takes the Fallback, which is the modern form, so
 * reporting the Choice told the user a "newer PowerPoint feature" had been
 * replaced when nothing was. A Choice with no Fallback, or one that also
 * requires any other unsupported token, still reports.
 */
const LEGACY_CHOICE_NAMESPACES: ReadonlySet<string> = new Set(['v']);

function isLegacyChoiceWithFallback(choice: XmlObject, block: XmlObject | undefined): boolean {
	if (!block?.['mc:Fallback']) {
		return false;
	}
	const tokens = String(choice['@_Requires'] ?? '')
		.trim()
		.split(/\s+/)
		.filter(Boolean);
	return tokens.length > 0 && tokens.every((token) => LEGACY_CHOICE_NAMESPACES.has(token));
}

function asArray(value: unknown): XmlObject[] {
	if (!value) {
		return [];
	}
	return (Array.isArray(value) ? value : [value]) as XmlObject[];
}
