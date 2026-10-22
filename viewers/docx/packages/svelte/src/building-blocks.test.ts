// @vitest-environment jsdom
import { expect, it } from 'vitest';
import { documentCanvas } from './building-blocks';
import { createBuildingBlockSession } from '../../bindings/src/building-blocks';
import { createDocument } from '../../core/src/index';
it('preserves unsaved edits when a Svelte action receives fresh props with the same mount callback', () => {
	const session = createBuildingBlockSession({ documentModel: createDocument(), canEdit: true });
	const host = document.createElement('div');
	document.body.append(host);
	const action = documentCanvas(host, session.canvasProps);
	try {
		const engine = session.handle.element;
		session.state.toolbarProps!.onBold();
		action.update({ ...session.canvasProps });
		expect(session.handle.element).toBe(engine);
		expect(session.state.toolbarProps?.bold).toBe(true);
	} finally {
		action.destroy();
		session.destroy();
		host.remove();
	}
	expect(host.childElementCount).toBe(0);
});
