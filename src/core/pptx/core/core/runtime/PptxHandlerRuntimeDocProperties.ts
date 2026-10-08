import { XmlObject } from '../../types';
import type {
	PptxActiveXControl,
	PptxAppProperties,
	PptxCoreProperties,
	PptxCustomProperty,
	PptxCustomerData,
	PptxTagCollection,
} from '../../types';
import {
	parseAppProperties as parseSharedAppProperties,
	parseCoreProperties as parseSharedCoreProperties,
	parseCustomPropertyTexts,
} from '../../../../opc/properties/index';
import { parseActiveXControlsFromSlide } from '../../utils/activex-parser';
import { resolveContentType } from '../../utils/customer-data-package';
import { safeResolveZipPath } from '../../utils/safe-path';
import {
	toPptxAppProperties,
	toPptxCoreProperties,
	toPptxCustomProperties,
} from '../../utils/document-properties-model';
import { discoverTagCollections } from '../../utils/tag-package';
import { PptxHandlerRuntime as PptxHandlerRuntimeBase } from './PptxHandlerRuntimeMediaData';

export class PptxHandlerRuntime extends PptxHandlerRuntimeBase {
	protected buildRelativeTargetPath(fromPartPath: string, toPartPath: string): string {
		const fromParts = fromPartPath.split('/');
		const toParts = toPartPath.split('/');
		// Remove file name from source part path.
		fromParts.pop();

		while (fromParts.length > 0 && toParts.length > 0 && fromParts[0] === toParts[0]) {
			fromParts.shift();
			toParts.shift();
		}

		const upSegments = Array.from<string>({ length: fromParts.length }).fill('..');
		return [...upSegments, ...toParts].join('/');
	}

	protected async setMasterThemeRelationship(masterPath: string, themePath: string): Promise<void> {
		const relsPath = masterPath.replace(
			/ppt\/slideMasters\/(slideMaster\d+)\.xml/,
			'ppt/slideMasters/_rels/$1.xml.rels',
		);
		const relsFile = this.zip.file(relsPath);
		if (!relsFile) {
			return;
		}

		const relsXml = await relsFile.async('string');
		const relsData = this.parser.parse(relsXml) as XmlObject;
		const relRoot = (relsData['Relationships'] || {}) as XmlObject;
		const relationships = this.ensureArray(relRoot['Relationship']) as XmlObject[];
		const themeRel = relationships.find((rel) => String(rel['@_Type'] || '').includes('/theme'));
		if (!themeRel) {
			return;
		}

		themeRel['@_Target'] = this.buildRelativeTargetPath(masterPath, themePath);
		relRoot['Relationship'] = relationships;
		relsData['Relationships'] = relRoot;
		this.zip.file(relsPath, this.builder.build(relsData));
	}

	public async setPresentationTheme(themePath: string, applyToAllMasters = true): Promise<void> {
		const normalizedThemePath = themePath.trim().replace(/\\/g, '/');
		if (!normalizedThemePath.startsWith('ppt/theme/')) {
			return;
		}
		const masterFiles = this.zip.file(/^ppt\/slideMasters\/slideMaster\d+\.xml$/);
		if (!masterFiles || masterFiles.length === 0) {
			return;
		}

		const targetMasters = applyToAllMasters ? masterFiles : [masterFiles[0]];
		await Promise.all(
			targetMasters.map(async (masterFile) => {
				await this.setMasterThemeRelationship(masterFile.name, normalizedThemePath);
			}),
		);
	}

	/** The text of a package part, or `undefined` when the package has no such part. */
	private async readPartText(path: string): Promise<string | undefined> {
		return this.zip.file(path)?.async('string');
	}

	/**
	 * Parse extended (application) properties from `docProps/app.xml`.
	 */
	protected async parseAppProperties(): Promise<PptxAppProperties | undefined> {
		try {
			const xml = await this.readPartText('docProps/app.xml');
			return xml === undefined ? undefined : toPptxAppProperties(parseSharedAppProperties(xml));
		} catch (e) {
			console.warn('Failed to parse app properties:', e);
			return undefined;
		}
	}

	/**
	 * Parse core document properties from `docProps/core.xml`.
	 */
	protected async parseCoreProperties(): Promise<PptxCoreProperties | undefined> {
		try {
			const xml = await this.readPartText('docProps/core.xml');
			return xml === undefined ? undefined : toPptxCoreProperties(parseSharedCoreProperties(xml));
		} catch (e) {
			console.warn('Failed to parse core properties:', e);
			return undefined;
		}
	}

	/**
	 * Parse custom document properties from `docProps/custom.xml`.
	 */
	protected async parseCustomProperties(): Promise<PptxCustomProperty[]> {
		try {
			const xml = await this.readPartText('docProps/custom.xml');
			return toPptxCustomProperties(parseCustomPropertyTexts(xml));
		} catch (e) {
			console.warn('Failed to parse custom properties:', e);
			return [];
		}
	}

	/**
	 * Parse `p:custDataLst` entries from a given XML container node and resolve
	 * their relationship targets + data content from the ZIP.
	 *
	 * @param containerNode - The XML node that may contain `p:custDataLst`.
	 * @param relsPath - Path to the `.rels` file for resolving relationship IDs.
	 * @param partPath - The owning part path (used to resolve relative targets).
	 */
	protected async parseCustDataLst(
		containerNode: XmlObject | undefined,
		relsPath: string,
		partPath: string,
	): Promise<PptxCustomerData[]> {
		if (!containerNode) {
			return [];
		}
		const custDataLst = containerNode['p:custDataLst'] as XmlObject | undefined;
		if (!custDataLst) {
			return [];
		}

		const custDataEntries = this.ensureArray(custDataLst['p:custData']) as XmlObject[];
		if (custDataEntries.length === 0) {
			return [];
		}

		// Load the relationships file to resolve r:id targets
		const relsFile = this.zip.file(relsPath);
		if (!relsFile) {
			return [];
		}

		const relsXml = await relsFile.async('string');
		const relsData = this.parser.parse(relsXml) as XmlObject;
		const relRoot = (relsData['Relationships'] || {}) as XmlObject;
		const relationships = this.ensureArray(relRoot['Relationship']) as XmlObject[];

		const relMap = new Map<string, string>();
		for (const rel of relationships) {
			const id = String(rel['@_Id'] || '').trim();
			const target = String(rel['@_Target'] || '').trim();
			const type = String(rel['@_Type'] || '').trim();
			if (id && target && type.endsWith('/relationships/customXml')) {
				relMap.set(id, target);
			}
		}

		const results: PptxCustomerData[] = [];
		for (const entry of custDataEntries) {
			const relId = String(entry['@_r:id'] || '').trim();
			if (!relId) {
				continue;
			}

			const target = relMap.get(relId);
			if (!target) {
				continue;
			}

			const base = partPath.slice(0, partPath.lastIndexOf('/'));
			const resolvedPath = safeResolveZipPath(base, target);
			if (!resolvedPath) {
				continue;
			}

			let data: string | undefined;
			try {
				const file = this.zip.file(resolvedPath);
				if (file) {
					data = await file.async('string');
				}
			} catch {
				// Non-critical — data may not be resolvable
			}

			const contentType = await resolveContentType(
				this.zip,
				{ parse: (xml) => this.parser.parse(xml) as XmlObject },
				resolvedPath,
			);
			results.push({ id: resolvedPath, relId, data, contentType, rawXml: entry });
		}

		return results;
	}

	/**
	 * Parse presentation-level customer data from `p:custDataLst` in
	 * `presentation.xml`.
	 */
	protected async parsePresentationCustomerData(): Promise<PptxCustomerData[]> {
		try {
			const presentation = this.presentationData?.['p:presentation'] as XmlObject | undefined;
			return await this.parseCustDataLst(
				presentation,
				'ppt/_rels/presentation.xml.rels',
				'ppt/presentation.xml',
			);
		} catch (e) {
			console.warn('Failed to parse presentation customer data:', e);
			return [];
		}
	}

	/**
	 * Parse slide-level customer data from `p:custDataLst` within `p:cSld`.
	 */
	protected async parseSlideCustomerData(
		slideXml: XmlObject,
		slidePath: string,
	): Promise<PptxCustomerData[]> {
		try {
			const sld = slideXml['p:sld'] as XmlObject | undefined;
			const cSld = sld?.['p:cSld'] as XmlObject | undefined;
			const relsPath = `${slidePath.replace('slides/', 'slides/_rels/')}.rels`;
			return await this.parseCustDataLst(cSld, relsPath, slidePath);
		} catch (e) {
			console.warn(`Failed to parse slide customer data for ${slidePath}:`, e);
			return [];
		}
	}

	/**
	 * Parse `p:controls > p:control` entries from a slide's `p:cSld`.
	 */
	protected parseSlideActiveXControls(slideXml: XmlObject): PptxActiveXControl[] {
		return parseActiveXControlsFromSlide(slideXml);
	}

	/**
	 * Read the package thumbnail image from `docProps/thumbnail.jpeg`.
	 *
	 * The thumbnail is a binary image stored in the OPC package and is
	 * preserved as raw bytes for lossless round-trip. Checks common
	 * extension variants (.jpeg, .jpg, .png, .emf).
	 */
	protected async parseThumbnail(): Promise<Uint8Array | null> {
		const candidates = [
			'docProps/thumbnail.jpeg',
			'docProps/thumbnail.jpg',
			'docProps/thumbnail.png',
			'docProps/thumbnail.emf',
		];
		for (const path of candidates) {
			const file = this.zip.file(path);
			if (file) {
				try {
					return await file.async('uint8array');
				} catch {
					// Non-critical — skip if unreadable
				}
			}
		}
		return null;
	}

	/**
	 * Parse all tag collections from `ppt/tags/tag*.xml`.
	 */
	protected async parseTags(): Promise<PptxTagCollection[]> {
		try {
			return await discoverTagCollections(this.zip, {
				parse: (xml) => this.parser.parse(xml) as XmlObject,
			});
		} catch (e) {
			console.warn('Failed to parse tags:', e);
			return [];
		}
	}
}
