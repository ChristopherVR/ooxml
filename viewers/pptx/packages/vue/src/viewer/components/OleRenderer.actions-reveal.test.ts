import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

import { expect, it } from 'vitest';

// The Download / Open bar is hidden until the object is hovered or a control in
// it is focused (as in React, Angular and vanilla): PowerPoint paints the
// embedded object alone, and a permanent bar changed every OLE slide's still.
it('hides the OLE action bar until hover or focus', () => {
	const source = readFileSync(
		resolve(process.cwd(), 'src/viewer/components/OleRenderer.vue'),
		'utf8',
	);
	const bar = /\.pptx-vue-ole-actions \{[^}]*\}/u.exec(source)?.[0] ?? '';
	expect(bar).toMatch(/opacity:\s*0;/u);
	expect(source).toMatch(
		/\.pptx-vue-ole:hover \.pptx-vue-ole-actions,\s*\.pptx-vue-ole-actions:focus-within \{\s*opacity:\s*1;/u,
	);
});
