// The pixel size a picture `data:` URL declares in its header (PNG, GIF, BMP, JPEG), which chart
// picture fills tile and stack at.

function decodeBase64Head(url: string, bytes: number): Uint8Array | undefined {
	const match = /^data:[^;,]*;base64,/u.exec(url);
	if (!match) return undefined;
	const chars = Math.ceil((bytes * 4) / 3 / 4) * 4;
	try {
		const binary = atob(url.slice(match[0].length, match[0].length + chars).replace(/\s/gu, ''));
		return Uint8Array.from(binary, (char) => char.charCodeAt(0));
	} catch {
		return undefined;
	}
}

/** The pixel size a PNG, GIF, BMP or JPEG `data:` URL declares in its header, if readable. */
export function chartPictureSize(url: string): { width: number; height: number } | undefined {
	const head = decodeBase64Head(url, 64 * 1024);
	if (!head || head.length < 10) return undefined;
	const view = new DataView(head.buffer, head.byteOffset, head.byteLength);
	if (head[0] === 0x89 && head[1] === 0x50 && head[2] === 0x4e && head[3] === 0x47)
		return head.length < 24 ? undefined : { width: view.getUint32(16), height: view.getUint32(20) };
	if (head[0] === 0x47 && head[1] === 0x49 && head[2] === 0x46)
		return { width: view.getUint16(6, true), height: view.getUint16(8, true) };
	if (head[0] === 0x42 && head[1] === 0x4d && head.length >= 26)
		return { width: Math.abs(view.getInt32(18, true)), height: Math.abs(view.getInt32(22, true)) };
	if (head[0] === 0xff && head[1] === 0xd8) {
		let at = 2;
		while (at + 9 < head.length) {
			if (head[at] !== 0xff) return undefined;
			const marker = head[at + 1]!;
			const length = view.getUint16(at + 2);
			if (marker >= 0xc0 && marker <= 0xcf && ![0xc4, 0xc8, 0xcc].includes(marker))
				return { width: view.getUint16(at + 7), height: view.getUint16(at + 5) };
			at += 2 + length;
		}
	}
	return undefined;
}
