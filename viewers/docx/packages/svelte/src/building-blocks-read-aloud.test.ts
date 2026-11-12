// @vitest-environment jsdom
import { afterEach, expect, it, vi } from 'vitest';
import { documentCanvas } from './building-blocks';
import { createBuildingBlockSession } from '../../bindings/src/building-blocks';
import { stubSpeech } from '../../bindings/src/speech-stub';
import { createDocument } from '../../core/src/index';

afterEach(() => {
	document.body.replaceChildren();
	vi.unstubAllGlobals();
});

it('exposes Read Aloud state through the Svelte canvas action and stops on destroy', () => {
	const speech = stubSpeech();
	const model = createDocument();
	model.blocks = [{ type: 'paragraph', id: 'p', runs: [{ text: 'Svelte words' }] }];
	const session = createBuildingBlockSession({ documentModel: model });
	const host = document.createElement('div');
	document.body.append(host);
	const action = documentCanvas(host, session.canvasProps);
	try {
		expect(session.state.toolbarProps).toMatchObject({ canReadAloud: true, readingAloud: false });
		expect(session.state.toolbarProps!.onReadAloud()).toBe('started');
		expect(speech.spoken).toEqual(['Svelte words']);
		expect(session.state.toolbarProps?.readingAloud).toBe(true);
		action.destroy();
		expect(speechSynthesis.cancel).toHaveBeenCalled();
	} finally {
		session.destroy();
	}
});
