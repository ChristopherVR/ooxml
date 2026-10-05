// Original one-pixel black images generated with @napi-rs/canvas 1.0.10.
const fixtures = {
	'image/png':
		'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAABHNCSVQICAgIfAhkiAAAAAFzUkdCAK7OHOkAAAANSURBVAiZY2BgYPgPAAEEAQB9ssjfAAAAAElFTkSuQmCC',
	'image/jpeg':
		'/9j/4AAQSkZJRgABAQAAAQABAAD/4gHYSUNDX1BST0ZJTEUAAQEAAAHIAAAAAAQwAABtbnRyUkdCIFhZWiAH4AABAAEAAAAAAABhY3NwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAQAA9tYAAQAAAADTLQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAlkZXNjAAAA8AAAACRyWFlaAAABFAAAABRnWFlaAAABKAAAABRiWFlaAAABPAAAABR3dHB0AAABUAAAABRyVFJDAAABZAAAAChnVFJDAAABZAAAAChiVFJDAAABZAAAAChjcHJ0AAABjAAAADxtbHVjAAAAAAAAAAEAAAAMZW5VUwAAAAgAAAAcAHMAUgBHAEJYWVogAAAAAAAAb6IAADj1AAADkFhZWiAAAAAAAABimQAAt4UAABjaWFlaIAAAAAAAACSgAAAPhAAAts9YWVogAAAAAAAA9tYAAQAAAADTLXBhcmEAAAAAAAQAAAACZmYAAPKnAAANWQAAE9AAAApbAAAAAAAAAABtbHVjAAAAAAAAAAEAAAAMZW5VUwAAACAAAAAcAEcAbwBvAGcAbABlACAASQBuAGMALgAgADIAMAAxADb/2wBDAAEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQH/2wBDAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQH/wAARCAABAAEDASIAAhEBAxEB/8QAFQABAQAAAAAAAAAAAAAAAAAAAAv/xAAUEAEAAAAAAAAAAAAAAAAAAAAA/8QAFAEBAAAAAAAAAAAAAAAAAAAAAP/EABQRAQAAAAAAAAAAAAAAAAAAAAD/2gAMAwEAAhEDEQA/AJ/4AP/Z',
	'image/gif': 'R0lGODlhAQABAIAAAAAAAAAAACH5BAQAAAAALAAAAAABAAEAgAAAAAAAAAICRAEAOw==',
};

export function rasterFixture(mimeType = 'image/png', minBytes = 0) {
	let bytes = Uint8Array.from(Buffer.from(fixtures[mimeType], 'base64'));
	if (mimeType === 'image/png' && minBytes > bytes.length) {
		const length = Math.max(0, minBytes - bytes.length - 12);
		const chunk = new Uint8Array(length + 12);
		const view = new DataView(chunk.buffer);
		view.setUint32(0, length);
		chunk.set([110, 112, 65, 68], 4); // private ancillary npAD padding, ignored by image decoders
		let crc = 0xffffffff;
		for (const value of chunk.subarray(4, chunk.length - 4)) {
			crc ^= value;
			for (let i = 0; i < 8; i++) crc = crc & 1 ? 0xedb88320 ^ (crc >>> 1) : crc >>> 1;
		}
		view.setUint32(chunk.length - 4, (crc ^ 0xffffffff) >>> 0);
		const result = new Uint8Array(bytes.length + chunk.length);
		const at = bytes.length - 12; // insert before IEND
		result.set(bytes.subarray(0, at));
		result.set(chunk, at);
		result.set(bytes.subarray(at), at + chunk.length);
		bytes = result;
	}
	return { mimeType, bytes, pixelWidth: 1, pixelHeight: 1 };
}
