import { mkdir, readFile, writeFile } from 'node:fs/promises';
import { resolve } from 'node:path';

import { PptxHandler, setAfterAnimation } from '../../../src/core/dist/pptx/index.mjs';

const output = resolve(process.argv[2] ?? 'after-effect-clear-native');
await mkdir(output, { recursive: true });
const source = await readFile(
	process.argv[3] ?? new URL('../fixtures/animation-after-effect.pptx', import.meta.url),
);
await writeFile(resolve(output, 'source.pptx'), source);
for (const action of [
	'dimToColor',
	'dimScheme',
	'hideAfterAnimation',
	'hideOnNextClick',
] as const) {
	const handler = new PptxHandler();
	try {
		const data = await handler.load(
			source.buffer.slice(source.byteOffset, source.byteOffset + source.byteLength),
		);
		const slide = data.slides[0]!;
		const target = slide.nativeAnimations!.find((animation) =>
			action === 'dimScheme'
				? animation.target?.type === 'shape' && animation.target.shapeId === '3'
				: animation.afterAnimationAction === action,
		)!;
		slide.animations = setAfterAnimation(slide.animations ?? [], target.targetId!, 'none');
		await writeFile(resolve(output, `${action}-cleared.pptx`), await handler.save(data.slides));
	} finally {
		handler.dispose();
	}
}
