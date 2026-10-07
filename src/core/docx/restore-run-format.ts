import type { TextRun } from './model.js';
import { DIRECT_RUN_PROPERTY_KEYS } from './run-formatting.js';
import { parseDirectRunProperties } from './run-properties.js';
import type { XmlElement } from './xml.js';
import { parsePropertiesSnapshot } from './revision-properties.js';

export function parseRunPropertiesSnapshot(xml: string): XmlElement {
	return parsePropertiesSnapshot(xml, 'rPr');
}

/** Restores modeled and opaque properties without changing text, anchors or inline content. */
export function restoreRunFormatting(run: TextRun): void {
	const xml = run.revision?.previousRunPropertiesXml;
	if (!xml)
		throw new Error(
			'Cannot reject a formatting revision without its prior run-properties snapshot.',
		);
	if (run.image) throw new Error('Rejecting image formatting revisions is not supported yet.');
	const previous = parseDirectRunProperties(parseRunPropertiesSnapshot(xml));
	for (const key of DIRECT_RUN_PROPERTY_KEYS) delete run[key];
	Object.assign(run, previous);
	run.restoredRunPropertiesXml = xml;
	delete run.revision;
}
