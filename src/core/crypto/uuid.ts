/** Office rule/series identity generation; this helper does not generate encryption secrets. */
function randomBytes(length: number): Uint8Array {
	const cryptoObj = globalThis.crypto;
	if (cryptoObj?.getRandomValues) return cryptoObj.getRandomValues(new Uint8Array(length));
	const bytes = new Uint8Array(length);
	for (let i = 0; i < length; i++) bytes[i] = Math.floor(Math.random() * 256);
	return bytes;
}

function randomHex(length: number): string {
	return Array.from(randomBytes(Math.ceil(length / 2)), (b) => b.toString(16).padStart(2, '0'))
		.join('')
		.slice(0, length);
}

/** A fresh braced uppercase GUID, including older browsers without crypto.randomUUID. */
export function createOfficeGuid(): string {
	const variantNibble = (8 + ((randomBytes(1)[0] ?? 0) % 4)).toString(16);
	const uuid =
		globalThis.crypto?.randomUUID?.() ??
		`${randomHex(8)}-${randomHex(4)}-4${randomHex(3)}-${variantNibble}${randomHex(3)}-${randomHex(12)}`;
	return `{${uuid.toUpperCase()}}`;
}
