import type { TextRun } from './model';
import { DIRECT_RUN_PROPERTY_KEYS } from './run-formatting';
import { parseDirectRunProperties } from './run-properties';
import type { XmlElement } from './xml';
import { parsePropertiesSnapshot } from './revision-properties';
import { runPropertiesHaveUnknownContent } from './write-run-validation';
import { buildXml } from './xml';

export function parseRunPropertiesSnapshot(xml: string): XmlElement {
	return parsePropertiesSnapshot(xml, 'rPr');
}

/** Restores modeled and opaque properties without changing text, anchors or inline content. */
export function restoreRunFormatting(run: TextRun): void {
	const nested = run.formatRevision;
	const xml = (nested ?? run.revision)?.previousRunPropertiesXml;
	if (!xml)
		throw new Error(
			'Cannot reject a formatting revision without its prior run-properties snapshot.',
		);
	const snapshot = parseRunPropertiesSnapshot(xml);
	const previous = parseDirectRunProperties(snapshot);
	for (const key of DIRECT_RUN_PROPERTY_KEYS) delete run[key];
	Object.assign(run, previous);
	run.restoredRunPropertiesXml = xml;
	if (runPropertiesHaveUnknownContent(snapshot)) run.sourceRunPropertiesXml = buildXml(snapshot);
	else delete run.sourceRunPropertiesXml;
	if (nested) delete run.formatRevision;
	else delete run.revision;
}
