// Synthetic inputs, written for this audit. No user document bytes are included.
// Scenarios: https://github.com/ChristopherVR/emf-converter/issues/{19,20,21}.
export function record(type, words = []) {
	const out = Buffer.alloc(8 + words.length * 4);
	out.writeUInt32LE(type, 0);
	out.writeUInt32LE(out.length, 4);
	words.forEach((word, index) => out.writeUInt32LE(word >>> 0, 8 + index * 4));
	return out;
}

export function emf(records, width = 300, height = 200) {
	const header = Buffer.alloc(108);
	header.writeUInt32LE(1, 0);
	header.writeUInt32LE(108, 4);
	header.writeInt32LE(width - 1, 16);
	header.writeInt32LE(height - 1, 20);
	header.writeInt32LE(Math.round((width * 2540) / 96), 32);
	header.writeInt32LE(Math.round((height * 2540) / 96), 36);
	header.writeUInt32LE(0x464d4520, 40);
	header.writeUInt32LE(0x10000, 44);
	header.writeUInt16LE(16, 56);
	[1920, 1080, 508, 286].forEach((n, i) => header.writeInt32LE(n, 72 + i * 4));
	header.writeInt32LE(508000, 100);
	header.writeInt32LE(286000, 104);
	const all = [header, ...records, record(14, [0, 16, 20])];
	header.writeUInt32LE(
		all.reduce((n, b) => n + b.length, 0),
		48,
	);
	header.writeUInt32LE(all.length, 52);
	return Buffer.concat(all);
}

export function plus(type, flags = 0, data = Buffer.alloc(0)) {
	const out = Buffer.alloc((12 + data.length + 3) & ~3);
	out.writeUInt16LE(type, 0);
	out.writeUInt16LE(flags, 2);
	out.writeUInt32LE(out.length, 4);
	out.writeUInt32LE(data.length, 8);
	data.copy(out, 12);
	return out;
}

export function comment(records) {
	const bytes = Buffer.concat([Buffer.from('EMF+'), ...records]);
	const out = Buffer.alloc(12 + bytes.length);
	out.writeUInt32LE(70, 0);
	out.writeUInt32LE(out.length, 4);
	out.writeUInt32LE(bytes.length, 8);
	bytes.copy(out, 12);
	return out;
}

function fields(size, ints = [], floats = []) {
	const out = Buffer.alloc(size);
	ints.forEach(([o, n]) => out.writeUInt32LE(n >>> 0, o));
	floats.forEach(([o, n]) => out.writeFloatLE(n, o));
	return out;
}

function plusFile(records) {
	return emf([
		comment([
			plus(
				0x4001,
				1,
				fields(16, [
					[0, 0xdbc01002],
					[8, 96],
					[12, 96],
				]),
			),
			...records,
		]),
		comment([plus(0x4002)]),
	]);
}

function fill(argb, width = 300) {
	return plus(
		0x400a,
		0x8000,
		fields(
			24,
			[
				[0, argb],
				[4, 1],
			],
			[
				[16, width],
				[20, 200],
			],
		),
	);
}

function font() {
	const data = fields(
		34,
		[
			[0, 0xdbc01002],
			[8, 2],
			[20, 5],
		],
		[[4, 36]],
	);
	Buffer.from('Arial', 'utf16le').copy(data, 24);
	return plus(0x4008, 0x0601, data);
}

function driver(text, flags = 1, withMatrix = false) {
	const n = text.length;
	const data = fields(16 + n * 10 + (withMatrix ? 24 : 0), [
		[0, 0xff000000],
		[4, flags],
		[8, +withMatrix],
		[12, n],
	]);
	Buffer.from(text, 'utf16le').copy(data, 16);
	for (let i = 0; i < n; i++) {
		data.writeFloatLE(110 + i * 26, 16 + n * 2 + i * 8);
		data.writeFloatLE(110, 20 + n * 2 + i * 8);
	}
	if (withMatrix) {
		[1, 0, 0, 1, 40, 0].forEach((v, i) => data.writeFloatLE(v, 16 + n * 10 + i * 4));
	}
	return plus(0x4036, 0x8001, data);
}

function drawString(text) {
	const data = fields(
		28 + text.length * 2,
		[
			[0, 0xff000000],
			[4, 2],
			[8, text.length],
		],
		[
			[12, 100],
			[16, 70],
			[20, 190],
			[24, 60],
		],
	);
	Buffer.from(text, 'utf16le').copy(data, 28);
	return plus(0x401c, 0x8001, data);
}

function clipCase(childCount, narrow = true) {
	const region = fields(
		28,
		[
			[0, 0xdbc01002],
			[4, childCount],
			[8, 0x10000000],
		],
		[
			[20, 300],
			[24, 200],
		],
	);
	return plusFile([
		...(narrow
			? [
					plus(
						0x4032,
						0,
						fields(
							16,
							[],
							[
								[8, 75],
								[12, 200],
							],
						),
					),
				]
			: []),
		plus(0x4008, 0x0403, region),
		plus(0x4034, 3),
		fill(0xffff0000),
	]);
}

export function syntheticCases() {
	const paint = [record(39, [1, 0, 0xff, 0]), record(37, [1]), record(43, [0, 0, 300, 200])];
	const extents = [record(9, [900, 600]), record(11, [300, 200])];
	const paddedOdd = plusFile([fill(0xff0000ff, 80), font(), driver('HALLO')]);
	for (let p = 0; p + 12 <= paddedOdd.length; p += 4) {
		if (paddedOdd.readUInt16LE(p) === 0x4036 && paddedOdd.readUInt16LE(p + 2) === 0x8001) {
			paddedOdd.writeUInt32LE(paddedOdd.readUInt32LE(p + 4) - 12, p + 8);
		}
	}
	return {
		'plain-rectangle': emf(paint),
		'mm-text-extents': emf([...extents, ...paint]),
		'anisotropic-extents': emf([record(17, [8]), ...extents, ...paint]),
		'draw-string-control': plusFile([
			fill(0xff0000ff, 80),
			font(),
			plus(
				0x4008,
				0x0702,
				fields(60, [
					[0, 0xdbc01002],
					[8, 0x409],
					[24, 0x409],
				]),
			),
			drawString('HALLO'),
		]),
		'driver-odd-count': plusFile([fill(0xff0000ff, 80), font(), driver('HALLO')]),
		'driver-odd-padded-data': paddedOdd,
		'driver-even-count': plusFile([fill(0xff0000ff, 80), font(), driver('HALO')]),
		'driver-even-matrix': plusFile([fill(0xff0000ff, 80), font(), driver('HALO', 1, true)]),
		'clip-replace-leaf': clipCase(0),
		'clip-leaf-control': clipCase(0, false),
		// Deliberately wrong count diagnoses the converter's zero-count rejection.
		'clip-replace-wrong-count': clipCase(1),
	};
}
