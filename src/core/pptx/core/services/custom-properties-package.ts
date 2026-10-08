/**
 * The package entries that go with `docProps/custom.xml`: its `[Content_Types].xml` override
 * and the package-root relationship, added when the part is written and removed with it.
 *
 * @module custom-properties-package
 */

import type { XMLBuilder, XMLParser } from 'fast-xml-parser';
import type JSZip from 'jszip';

import { PROPERTY_CONTENT_TYPES } from '../../../opc/properties/index';
import { RELATIONSHIP_TYPES } from '../../../opc/relationship-types';
import type { XmlObject } from '../types';

export interface CustomPropertiesPackageContext {
	zip: JSZip;
	parser: XMLParser;
	builder: XMLBuilder;
}

/**
 * Ensure `[Content_Types].xml` has an `Override` for `docProps/custom.xml`
 * and the root `_rels/.rels` references it (ECMA-376 §15.2.12.2 +
 * Part 2 §10.1.2.5). Without these, the package fails OPC validation
 * and Office strips the custom properties on next save.
 */
export async function ensureCustomPropertiesPackagingArtifacts(
	context: CustomPropertiesPackageContext,
): Promise<void> {
	// 1. [Content_Types].xml: add the Override if missing
	const ctFile = context.zip.file('[Content_Types].xml');
	if (ctFile) {
		try {
			const ctXml = await ctFile.async('string');
			const ctData = context.parser.parse(ctXml) as XmlObject;
			const types = ctData['Types'] as XmlObject | undefined;
			if (types) {
				const overrides = Array.isArray(types['Override'])
					? (types['Override'] as XmlObject[])
					: types['Override']
						? [types['Override'] as XmlObject]
						: [];
				const hasCustomOverride = overrides.some(
					(o) => String(o?.['@_PartName'] || '') === '/docProps/custom.xml',
				);
				if (!hasCustomOverride) {
					overrides.push({
						'@_PartName': '/docProps/custom.xml',
						'@_ContentType': PROPERTY_CONTENT_TYPES.custom,
					});
					types['Override'] = overrides.length === 1 ? overrides[0] : overrides;
					context.zip.file('[Content_Types].xml', context.builder.build(ctData));
				}
			}
		} catch (error) {
			console.warn('Failed to update [Content_Types].xml for custom properties:', error);
		}
	}

	// 2. _rels/.rels: add the custom-properties relationship if missing
	const relsFile = context.zip.file('_rels/.rels');
	if (relsFile) {
		try {
			const relsXml = await relsFile.async('string');
			const relsData = context.parser.parse(relsXml) as XmlObject;
			const relationships = relsData['Relationships'] as XmlObject | undefined;
			if (relationships) {
				const rels = Array.isArray(relationships['Relationship'])
					? (relationships['Relationship'] as XmlObject[])
					: relationships['Relationship']
						? [relationships['Relationship'] as XmlObject]
						: [];
				const hasCustomRel = rels.some(
					(r) => String(r?.['@_Type'] || '') === RELATIONSHIP_TYPES.customProperties,
				);
				if (!hasCustomRel) {
					// Compute next free rId
					let maxId = 0;
					for (const rel of rels) {
						const id = String(rel?.['@_Id'] || '');
						const num = Number.parseInt(id.replace(/^rId/, ''), 10);
						if (Number.isFinite(num) && num > maxId) {
							maxId = num;
						}
					}
					rels.push({
						'@_Id': `rId${maxId + 1}`,
						'@_Type': RELATIONSHIP_TYPES.customProperties,
						'@_Target': 'docProps/custom.xml',
					});
					relationships['Relationship'] = rels;
					context.zip.file('_rels/.rels', context.builder.build(relsData));
				}
			}
		} catch (error) {
			console.warn('Failed to update _rels/.rels for custom properties:', error);
		}
	}
}

/**
 * Remove the Override + root rel for `docProps/custom.xml` when the
 * caller has emptied custom properties so the package doesn't keep an
 * orphan content-type entry referencing a deleted part.
 */
export async function removeCustomPropertiesPackagingArtifacts(
	context: CustomPropertiesPackageContext,
): Promise<void> {
	const ctFile = context.zip.file('[Content_Types].xml');
	if (ctFile) {
		try {
			const ctXml = await ctFile.async('string');
			const ctData = context.parser.parse(ctXml) as XmlObject;
			const types = ctData['Types'] as XmlObject | undefined;
			if (types) {
				const overrides = Array.isArray(types['Override'])
					? (types['Override'] as XmlObject[])
					: types['Override']
						? [types['Override'] as XmlObject]
						: [];
				const filtered = overrides.filter(
					(o) => String(o?.['@_PartName'] || '') !== '/docProps/custom.xml',
				);
				if (filtered.length !== overrides.length) {
					types['Override'] = filtered.length === 1 ? filtered[0] : filtered;
					context.zip.file('[Content_Types].xml', context.builder.build(ctData));
				}
			}
		} catch {
			/* noop */
		}
	}

	const relsFile = context.zip.file('_rels/.rels');
	if (relsFile) {
		try {
			const relsXml = await relsFile.async('string');
			const relsData = context.parser.parse(relsXml) as XmlObject;
			const relationships = relsData['Relationships'] as XmlObject | undefined;
			if (relationships) {
				const rels = Array.isArray(relationships['Relationship'])
					? (relationships['Relationship'] as XmlObject[])
					: relationships['Relationship']
						? [relationships['Relationship'] as XmlObject]
						: [];
				const filtered = rels.filter(
					(r) => String(r?.['@_Type'] || '') !== RELATIONSHIP_TYPES.customProperties,
				);
				if (filtered.length !== rels.length) {
					relationships['Relationship'] = filtered;
					context.zip.file('_rels/.rels', context.builder.build(relsData));
				}
			}
		} catch {
			/* noop */
		}
	}
}
