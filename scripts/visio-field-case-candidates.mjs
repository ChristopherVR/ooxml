import { writeFileSync } from 'node:fs';

// Enumerate the runtime's caseable scalar values; no Office implementation is copied.
const characters = [];
for (let point = 0; point <= 0x10ffff; point++) {
	if (point >= 0xd800 && point <= 0xdfff) continue;
	const character = String.fromCodePoint(point);
	if (character.toUpperCase() !== character || character.toLowerCase() !== character)
		characters.push(character);
}
writeFileSync(
	process.argv[2],
	JSON.stringify({
		unicode: process.versions.unicode,
		inputs: Array.from({ length: Math.ceil(characters.length / 200) }, (_, index) =>
			characters.slice(index * 200, (index + 1) * 200).join(''),
		),
	}),
);
