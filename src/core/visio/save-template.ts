import JSZip from 'jszip';
import { buildXml, elements, parseXml } from '../xml/index';
import { writeEditedPackage } from './edit-package';
import { operationStep, withOperationMemory } from './memory-budget';
import { VisioPackage } from './package';
import { DEFAULTS, decodePath, fail, type VisioPackageLimits } from './package-common';
import {
	DRAWING_MACRO_MAIN,
	DRAWING_MAIN_TYPES,
	DRAWING_TO_TEMPLATE,
	TEMPLATE_MACRO_MAIN,
	stripVisioSignatures,
} from './package-kind';
import { related, visioXml } from './parts';
import { attribute } from './sheet';

export interface ConvertVsdxToTemplateOptions {
	limits?: Partial<VisioPackageLimits>;
	maxOutputBytes?: number;
}

const CONTENT_TYPES_NS = 'http://schemas.openxmlformats.org/package/2006/content-types';

function positive(value: number): number {
	if (!Number.isSafeInteger(value) || value <= 0)
		fail('INVALID_LIMITS', 'Template conversion limits must be positive safe integers.');
	return value;
}

/**
 * Convert a Visio drawing package (`.vsdx` or `.vsdm`) into a Visio template package (`.vstx` or
 * `.vstm`), keeping all pages, masters, styles, themes, relationships and (for macro-enabled
 * packages) the VBA project byte-for-byte. Digital signatures are stripped because rewriting
 * `[Content_Types].xml` invalidates them. Already-template packages without signatures are
 * returned unchanged after validation.
 */
export async function convertVsdxToTemplate(
	input: Uint8Array | ArrayBuffer,
	given: ConvertVsdxToTemplateOptions = {},
): Promise<Uint8Array> {
	const options: ConvertVsdxToTemplateOptions = {
		...given,
		limits: withOperationMemory(given.limits, input.byteLength, 'edit'),
	};
	return operationStep(options.limits, () => convertVsdxToTemplateWith(input, options));
}

async function convertVsdxToTemplateWith(
	input: Uint8Array | ArrayBuffer,
	options: ConvertVsdxToTemplateOptions,
): Promise<Uint8Array> {
	const limits = { ...DEFAULTS, ...options.limits };
	for (const value of Object.values(limits)) positive(value);
	const maxOutput = positive(options.maxOutputBytes ?? limits.maxInputBytes);
	const deadline = Date.now() + limits.maxRuntimeMs;
	const check = () => {
		if (Date.now() >= deadline)
			fail('LIMIT_RUNTIME', 'Visio template conversion deadline exceeded.');
	};

	const source = new Uint8Array(input instanceof Uint8Array ? input : new Uint8Array(input));
	if (source.byteLength > limits.maxInputBytes) fail('LIMIT_INPUT', 'ZIP input exceeds limit.');

	const pkg = await VisioPackage.open(source, limits, 2);
	const types = await pkg.readXml('[Content_Types].xml', 'Types');
	if (types.namespaceURI !== CONTENT_TYPES_NS)
		fail('INVALID_CONTENT_TYPES', 'Invalid content type namespace.');
	for (const node of elements(types))
		if (
			node.namespaceURI !== types.namespaceURI ||
			!['Default', 'Override'].includes(node.localName)
		)
			fail('INVALID_CONTENT_TYPES', 'Invalid content type element.');

	const documentPart = await related(pkg, '', 'document');
	if (!documentPart) fail('INVALID_DOCUMENT', 'Missing Visio document part.');
	const overrides = elements(types).filter(
		(node) => node.localName === 'Override' && attribute(node, 'PartName') === `/${documentPart}`,
	);
	if (overrides.length !== 1)
		fail(
			'UNSUPPORTED_FORMAT',
			'Saving as a template supports only Visio drawing and template packages.',
		);
	const mainType = overrides[0]!.getAttribute('ContentType') ?? '';
	if (!DRAWING_MAIN_TYPES.has(mainType))
		fail(
			'UNSUPPORTED_FORMAT',
			'Saving as a template supports only Visio drawing and template packages.',
		);
	const macroEnabled = mainType === DRAWING_MACRO_MAIN || mainType === TEMPLATE_MACRO_MAIN;
	const targetType = DRAWING_TO_TEMPLATE.get(mainType) ?? mainType;
	const needsTypeChange = targetType !== mainType;

	const zip = new JSZip();
	for (const path of pkg.paths()) {
		check();
		const decoded = decodePath(path);
		if (!macroEnabled && /(^|\/)vbaProject\.bin$/i.test(decoded))
			fail('UNSUPPORTED_EDIT_PACKAGE', 'The package has a VBA project but is not macro-enabled.');
		if (path === '_rels/.rels' || /(^|\/)_rels\/[^/]+\.rels$/.test(path)) {
			const relSource =
				path === '_rels/.rels' ? '' : path.replace(/(^|\/)_rels\/([^/]+)\.rels$/, '$1$2');
			for (const rel of (await pkg.relationships(relSource)).values())
				if (!macroEnabled && /vbaProject/i.test(rel.type))
					fail(
						'UNSUPPORTED_EDIT_PACKAGE',
						'The package has a VBA project but is not macro-enabled.',
					);
		}
		zip.file(path, new Uint8Array(await pkg.readBytes(path)), { createFolders: false });
	}
	await visioXml(pkg, documentPart, 'VisioDocument');

	const signatureRemoved = await stripVisioSignatures(zip);
	if (!needsTypeChange && !signatureRemoved) {
		if (source.byteLength > maxOutput)
			fail('LIMIT_EDIT_OUTPUT', 'Saved package exceeds output limit.');
		return source;
	}

	if (needsTypeChange) {
		const typesFile = zip.file('[Content_Types].xml');
		if (!typesFile) fail('INVALID_CONTENT_TYPES', 'Missing [Content_Types].xml.');
		const doc = parseXml(await typesFile.async('string'));
		for (const node of elements(doc.documentElement)) {
			if (
				node.localName === 'Override' &&
				node.getAttribute('PartName') === `/${documentPart}` &&
				node.getAttribute('ContentType') === mainType
			) {
				node.setAttribute('ContentType', targetType);
			}
		}
		zip.file('[Content_Types].xml', buildXml(doc));
	}

	const parts = new Map<string, Uint8Array>();
	for (const [path, file] of Object.entries(zip.files)) {
		check();
		if (file.dir) continue;
		parts.set(path, await file.async('uint8array'));
	}
	return writeEditedPackage(parts, maxOutput, deadline, check);
}
