// @vitest-environment jsdom
import { afterEach, expect, it, vi } from 'vitest';
import { createBuildingBlockSession } from './building-blocks';
import { toolbarControls } from './controls';
import { stubSpeech } from './speech-stub';
import { createDocument } from '../../core/src/index';

afterEach(() => {
	document.body.replaceChildren();
	vi.unstubAllGlobals();
});

function model() {
	const doc = createDocument();
	doc.blocks = [{ type: 'paragraph', id: 'p', runs: [{ text: 'Shared session text' }] }];
	return doc;
}

it('publishes playback state, stops on unmount and hides the control without speech', () => {
	const speech = stubSpeech();
	const session = createBuildingBlockSession({ documentModel: model() });
	const host = document.createElement('div');
	document.body.append(host);
	try {
		session.canvasProps.mount(host);
		const props = session.state.toolbarProps!;
		expect(props).toMatchObject({ canReadAloud: true, readingAloud: false, readAloudIssue: null });
		expect(props.onReadAloud()).toBe('started');
		expect(speech.spoken).toEqual(['Shared session text']);
		expect(session.state.toolbarProps?.readingAloud).toBe(true);
		const control = toolbarControls(session.state.toolbarProps!)
			.flat()
			.find((item) => item.id === 'readAloud');
		expect(control).toMatchObject({ kind: 'toggle', pressed: true, disabled: false });
		session.canvasProps.mount(null);
		expect(speechSynthesis.cancel).toHaveBeenCalled();
		expect(session.state.toolbarProps).toBeNull();
		vi.stubGlobal('speechSynthesis', undefined);
		session.canvasProps.mount(host);
		const unsupported = session.state.toolbarProps!;
		expect(unsupported).toMatchObject({ canReadAloud: false, readAloudIssue: 'unsupported' });
		expect(
			toolbarControls(unsupported)
				.flat()
				.some((item) => item.id === 'readAloud'),
		).toBe(false);
	} finally {
		session.destroy();
	}
});
