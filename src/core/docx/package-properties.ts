// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// Reads the package's property parts into the shared `opc/properties` model, for DOCPROPERTY.
import type JSZip from 'jszip';
import {
	PROPERTY_PART_NAMES,
	parseAppProperties,
	parseCoreProperties,
	parseCustomProperties,
	type DocumentPropertySet,
} from '../opc/properties/index';

/** Reads one part, or nothing when it is missing or not well-formed: a load never fails on it. */
async function read<T>(
	zip: JSZip,
	path: string,
	parse: (xml: string) => T,
): Promise<T | undefined> {
	const file = zip.file(path);
	if (!file) return undefined;
	try {
		return parse(await file.async('string'));
	} catch {
		return undefined;
	}
}

/** The core, extended and custom properties of a package; undefined when it has none. */
export async function readPackageProperties(zip: JSZip): Promise<DocumentPropertySet | undefined> {
	const [core, app, custom] = await Promise.all([
		read(zip, PROPERTY_PART_NAMES.core, parseCoreProperties),
		read(zip, PROPERTY_PART_NAMES.app, parseAppProperties),
		read(zip, PROPERTY_PART_NAMES.custom, parseCustomProperties),
	]);
	if (!core && !app && !custom) return undefined;
	return { core: core ?? {}, app: app ?? {}, custom: custom ?? [] };
}
