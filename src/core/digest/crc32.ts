const crcTable = Uint32Array.from({ length: 256 }, (_, value) => {
	for (let i = 0; i < 8; i++) value = value & 1 ? 0xedb88320 ^ (value >>> 1) : value >>> 1;
	return value >>> 0;
});
export function crc32(bytes: Uint8Array, crc = 0xffffffff): number {
	for (const byte of bytes) crc = (crcTable[(crc ^ byte) & 255] ?? 0) ^ (crc >>> 8);
	return crc >>> 0;
}
