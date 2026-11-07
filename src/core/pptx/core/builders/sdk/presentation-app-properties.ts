/**
 * `docProps/app.xml` of a presentation made from scratch, with the
 * application properties a caller passes to `PptxHandler.create`.
 *
 * The values go through the same override and part writer as
 * `save(..., { appProperties })` (`applyAppOverrides`, `writeAppProperties`),
 * so a value set at create time and the same value set at save time are
 * written alike.
 *
 * @module sdk/presentation-app-properties
 */

import { parseAppProperties, writeAppProperties } from '../../../../opc/properties/index';
import type { PptxAppProperties } from '../../types/metadata';
import { applyAppOverrides } from '../../utils/document-properties-model';

/** The `docProps/app.xml` text fields a new presentation accepts. */
export type PresentationAppPropertiesInput = Pick<
	PptxAppProperties,
	'application' | 'appVersion' | 'presentationFormat' | 'company' | 'manager' | 'template'
>;

const INPUT_FIELDS: ReadonlySet<string> = new Set([
	'application',
	'appVersion',
	'presentationFormat',
	'company',
	'manager',
	'template',
] satisfies (keyof PresentationAppPropertiesInput)[]);

/**
 * `AppVersion` as Office writes it (`16.0000`): PowerPoint refuses a package
 * whose `AppVersion` has another form (`16`, `16.0`) as corrupt.
 */
const APP_VERSION = /^\d+\.\d{4}$/u;

function blankAppXml(slideCount: number): string {
	return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
  xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>pptx-viewer-sdk</Application>
  <Slides>${slideCount}</Slides>
  <ScaleCrop>false</ScaleCrop>
  <LinksUpToDate>false</LinksUpToDate>
  <SharedDoc>false</SharedDoc>
  <HyperlinksChanged>false</HyperlinksChanged>
  <AppVersion>1.0000</AppVersion>
</Properties>`;
}

/** Refuse a field this input does not take and a value that is not text. */
function validate(input: PresentationAppPropertiesInput): void {
	for (const [field, value] of Object.entries(input)) {
		if (!INPUT_FIELDS.has(field)) {
			throw new RangeError(
				`appProperties.${field} cannot be set when a presentation is created; ` +
					`it takes ${[...INPUT_FIELDS].join(', ')} (counts are computed on save)`,
			);
		}
		if (value !== undefined && typeof value !== 'string') {
			throw new RangeError(`appProperties.${field} must be a string`);
		}
	}
	const version = input.appVersion?.trim();
	if (version && !APP_VERSION.test(version)) {
		throw new RangeError(
			`appProperties.appVersion must be digits, a dot and four digits (e.g. "16.0000"), got "${input.appVersion}"`,
		);
	}
}

/**
 * The `docProps/app.xml` part of a new presentation with `slideCount`
 * slides. Without `input` it names `pptx-viewer-sdk` version `1.0000`; a
 * field of `input` replaces that value, and an empty string leaves the
 * element out, as the save option does.
 *
 * @throws RangeError for a field the input does not take, a value that is
 *   not a string, or an `appVersion` PowerPoint would refuse.
 */
export function newPresentationAppXml(
	slideCount: number,
	input?: PresentationAppPropertiesInput,
): string {
	const xml = blankAppXml(slideCount);
	if (input === undefined) {
		return xml;
	}
	validate(input);
	return writeAppProperties(applyAppOverrides(parseAppProperties(xml), input), xml);
}
