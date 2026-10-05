import JSZip from 'jszip';
import { parseXml } from '../xml/index.js';
import {
	DEFAULTS,
	VisioPackageError,
	fail,
	safePath,
	decodePath,
	crc32,
	type Entry,
	type VisioPackageLimits,
	type VisioPackageRelationship,
} from './package-common.js';
import { inspectZip } from './zip-validation.js';
import { inspectXml, inspectNamespaces } from './xml-validation.js';
export {
	VisioPackageError,
	type VisioPackageLimits,
	type VisioPackageRelationship,
} from './package-common.js';

function resolveTarget(source: string, target: string): string {
	const decoded = decodePath(target);
	if (!decoded || /[\\:#?%\u0000-\u001f\u007f]/.test(decoded) || decoded.startsWith('//'))
		fail('INVALID_RELATIONSHIP', 'Unsafe internal relationship target');
	const parts = decoded.startsWith('/') ? [] : decodePath(source).split('/').slice(0, -1);
	for (const part of decoded.split('/')) {
		if (!part || part === '.') continue;
		if (part === '..') {
			if (!parts.length) fail('INVALID_RELATIONSHIP', 'Relationship escapes package root');
			parts.pop();
		} else parts.push(part);
	}
	return safePath(parts.join('/'));
}

/** Read-only, bounded OPC access. External relationship targets are never fetched. */
export class VisioPackage {
	private readonly cache = new Map<string, Promise<Element>>();
	private readonly relationshipCache = new Map<
		string,
		Promise<Map<string, VisioPackageRelationship>>
	>();
	private readonly byteCache = new Map<string, Promise<Uint8Array>>();
	private readonly decodedNames = new Map<string, string>();
	private totalOutput = 0;
	private totalXmlNodes = 0;
	private constructor(
		private readonly zip: JSZip,
		private readonly entries: Map<string, Entry>,
		private readonly limits: VisioPackageLimits,
		private readonly deadline: number,
	) {
		for (const entry of entries.values())
			if (!entry.directory) this.decodedNames.set(decodePath(entry.name), entry.name);
	}
	private check = () => {
		if (Date.now() >= this.deadline)
			fail('LIMIT_RUNTIME', 'Visio package processing deadline exceeded');
	};

	static async open(
		input: Uint8Array | ArrayBuffer,
		options: Partial<VisioPackageLimits> = {},
	): Promise<VisioPackage> {
		const limits = { ...DEFAULTS, ...options };
		for (const value of Object.values(limits))
			if (!Number.isFinite(value) || value <= 0)
				fail('INVALID_LIMITS', 'Package limits must be finite positive numbers');
		const deadline = Date.now() + limits.maxRuntimeMs;
		const check = () => {
			if (Date.now() >= deadline)
				fail('LIMIT_RUNTIME', 'Visio package processing deadline exceeded');
		};
		const source = input instanceof Uint8Array ? input : new Uint8Array(input);
		if (source.byteLength > limits.maxInputBytes) fail('LIMIT_INPUT', 'ZIP input exceeds limit');
		// Own the bytes so the caller cannot mutate names or sizes after validation.
		const bytes = new Uint8Array(source);
		try {
			const entries = inspectZip(bytes, limits, check);
			const zip = await JSZip.loadAsync(bytes, { checkCRC32: false, createFolders: false });
			check();
			if (
				Object.keys(zip.files).length !== entries.size ||
				Object.keys(zip.files).some((name) => !entries.has(name))
			)
				fail('ZIP_MISMATCH', 'ZIP loader changed validated entry names');
			return new VisioPackage(zip, entries, limits, deadline);
		} catch (error) {
			if (error instanceof VisioPackageError) throw error;
			return fail(
				'INVALID_ZIP',
				`Invalid ZIP: ${error instanceof Error ? error.message : String(error)}`,
			);
		}
	}

	has(path: string): boolean {
		return this.entries.get(path)?.directory === false;
	}
	paths(): string[] {
		return [...this.entries.values()]
			.filter((entry) => !entry.directory)
			.map((entry) => entry.name);
	}

	async readXml(path: string, expectedRoot?: string): Promise<Element> {
		this.check();
		safePath(path);
		let pending = this.cache.get(path);
		if (!pending) {
			pending = this.loadXml(path);
			this.cache.set(path, pending);
		}
		const root = await pending;
		this.check();
		if (expectedRoot && root.localName !== expectedRoot)
			fail('INVALID_XML', `Expected ${expectedRoot} root in ${path}, found ${root.nodeName}`);
		return root;
	}

	/** Validated declared size, so resource-specific limits can reject before inflation. */
	getPartByteLength(path: string): number {
		this.check();
		safePath(path);
		const entry = this.entries.get(path);
		if (!entry || entry.directory) fail('MISSING_PART', `Missing package part: ${path}`);
		return entry.size;
	}
	/** Shared package-owned bytes. Treat the returned buffer as immutable. */
	async readBytes(path: string): Promise<Uint8Array> {
		this.check();
		safePath(path);
		let pending = this.byteCache.get(path);
		if (!pending) {
			pending = this.loadBytes(path);
			this.byteCache.set(path, pending);
		}
		const bytes = await pending;
		this.check();
		return bytes;
	}

	private async loadBytes(path: string): Promise<Uint8Array> {
		const entry = this.entries.get(path),
			file = this.zip.file(path);
		if (!entry || entry.directory || !file)
			return fail('MISSING_PART', `Missing package part: ${path}`);
		// JSZip's public stream API avoids its unbounded async()/accumulate() allocation.
		return new Promise<Uint8Array>((resolve, reject) => {
			const stream = (
				file as JSZip.JSZipObject & {
					internalStream(type: 'uint8array'): JSZip.JSZipStreamHelper<Uint8Array>;
				}
			).internalStream('uint8array');
			let chunks: Uint8Array[] = [],
				size = 0,
				crc = 0xffffffff,
				settled = false;
			const abort = (error: unknown) => {
				if (settled) return;
				settled = true;
				stream.pause();
				chunks = [];
				clearTimeout(timer);
				reject(
					error instanceof VisioPackageError
						? error
						: new VisioPackageError('INVALID_ZIP', `Cannot inflate ${path}: ${String(error)}`),
				);
			};
			const timer = setTimeout(
				() =>
					abort(
						new VisioPackageError('LIMIT_RUNTIME', 'Visio package processing deadline exceeded'),
					),
				Math.max(1, this.deadline - Date.now()),
			);
			stream
				.on('data', (chunk: Uint8Array) => {
					if (settled) return;
					try {
						this.check();
						size += chunk.length;
						this.totalOutput += chunk.length;
						if (size > this.limits.maxEntryBytes)
							fail('LIMIT_ENTRY', `Actual inflated entry exceeds limit: ${path}`);
						if (this.totalOutput > this.limits.maxTotalBytes)
							fail('LIMIT_TOTAL', 'Actual inflated ZIP total exceeds limit');
						if (size > Math.max(1, entry.compressed) * this.limits.maxCompressionRatio)
							fail('LIMIT_RATIO', `Actual ZIP compression ratio exceeds limit: ${path}`);
						if (size > entry.size)
							fail('ZIP_MISMATCH', `Actual ZIP entry exceeds declared size: ${path}`);
						crc = crc32(chunk, crc);
						chunks.push(chunk);
					} catch (error) {
						abort(error);
					}
				})
				.on('error', abort)
				.on('end', () => {
					if (settled) return;
					if (size !== entry.size || (crc ^ 0xffffffff) >>> 0 !== entry.crc) {
						abort(new VisioPackageError('ZIP_MISMATCH', `Inflated size or CRC mismatch: ${path}`));
						return;
					}
					settled = true;
					clearTimeout(timer);
					const result = new Uint8Array(size);
					let offset = 0;
					for (const chunk of chunks) {
						result.set(chunk, offset);
						offset += chunk.length;
					}
					chunks = [];
					resolve(result);
				})
				.resume();
		});
	}

	private async loadXml(path: string): Promise<Element> {
		const bytes = await this.readBytes(path);
		try {
			const encoding =
				(bytes[0] === 0xff && bytes[1] === 0xfe) || (bytes[0] === 0x3c && bytes[1] === 0)
					? 'utf-16le'
					: (bytes[0] === 0xfe && bytes[1] === 0xff) || (bytes[0] === 0 && bytes[1] === 0x3c)
						? 'utf-16be'
						: 'utf-8';
			const xml = new TextDecoder(encoding, { fatal: true }).decode(bytes);
			this.totalXmlNodes += inspectXml(xml, this.limits, this.check, encoding);
			if (this.totalXmlNodes > this.limits.maxTotalXmlNodes)
				fail('LIMIT_XML_TOTAL', 'Aggregate XML node limit exceeded');
			const root = parseXml(xml, { label: 'VSDX' }).documentElement;
			inspectNamespaces(root, this.check);
			this.check();
			return root;
		} catch (error) {
			if (error instanceof VisioPackageError) throw error;
			return fail(
				'INVALID_XML',
				`Invalid XML in ${path}: ${error instanceof Error ? error.message : String(error)}`,
			);
		}
	}

	async relationships(sourcePart: string): Promise<Map<string, VisioPackageRelationship>> {
		this.check();
		let pending = this.relationshipCache.get(sourcePart);
		if (!pending) {
			pending = this.loadRelationships(sourcePart);
			this.relationshipCache.set(sourcePart, pending);
		}
		return pending;
	}
	private async loadRelationships(
		sourcePart: string,
	): Promise<Map<string, VisioPackageRelationship>> {
		this.check();
		if (sourcePart) safePath(sourcePart);
		const slash = sourcePart.lastIndexOf('/');
		const path = sourcePart
			? `${sourcePart.slice(0, slash + 1)}_rels/${sourcePart.slice(slash + 1)}.rels`
			: '_rels/.rels';
		const result = new Map<string, VisioPackageRelationship>();
		if (!this.has(path)) return result;
		if (sourcePart && !this.has(sourcePart))
			fail('INVALID_RELATIONSHIP', `Missing relationship source: ${sourcePart}`);
		const root = await this.readXml(path, 'Relationships');
		if (root.namespaceURI !== 'http://schemas.openxmlformats.org/package/2006/relationships')
			fail('INVALID_RELATIONSHIP', 'Invalid relationships namespace');
		for (const node of Array.from(root.childNodes)) {
			this.check();
			if (node.nodeType !== 1) continue;
			const element = node as Element;
			if (element.localName !== 'Relationship' || element.namespaceURI !== root.namespaceURI)
				fail('INVALID_RELATIONSHIP', 'Invalid relationship element');
			const id = element.getAttribute('Id'),
				type = element.getAttribute('Type'),
				target = element.getAttribute('Target');
			const mode = element.getAttribute('TargetMode') || 'Internal';
			if (!id || !type || !target || (mode !== 'Internal' && mode !== 'External') || result.has(id))
				fail('INVALID_RELATIONSHIP', 'Missing, duplicate, or invalid relationship attributes');
			const resolved = mode === 'Internal' ? resolveTarget(sourcePart, target) : target;
			const actual = mode === 'Internal' ? this.decodedNames.get(resolved) : undefined;
			if (mode === 'Internal' && !actual)
				fail('INVALID_RELATIONSHIP', `Missing relationship target: ${resolved}`);
			result.set(id, { id, type, target: actual ?? resolved, mode });
		}
		return result;
	}
}
