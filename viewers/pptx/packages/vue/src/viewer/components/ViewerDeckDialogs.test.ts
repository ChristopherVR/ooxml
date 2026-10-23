import { mount } from '@vue/test-utils';
import { registerPptxWebControls } from 'ooxml-ui/pptx';
import { beforeAll, describe, expect, it, vi } from 'vitest';
import { ref } from 'vue';

import ViewerDeckDialogs from './ViewerDeckDialogs.vue';

vi.mock(import('vue-i18n'), () => ({
	useI18n: () => ({ t: (key: string) => key, locale: ref('en') }),
}));
beforeAll(registerPptxWebControls);

const closed = () => ref(false);

function mountDialogs() {
	const onInsertSmartArt = vi.fn();
	const insertDialogs = {
		showInsertSmartArt: ref(false),
		showEquationEditor: closed(),
		editingEquationOmml: ref(null),
		onInsertElement: vi.fn(),
		onInsertSmartArt,
		openEquationEditorForElement: vi.fn(),
		onApplyEquation: vi.fn(),
		closeEquationEditor: vi.fn(),
	};
	const wrapper = mount(ViewerDeckDialogs, {
		attachTo: document.body,
		props: {
			collaboration: {
				broadcastOpen: closed(),
				collabActive: closed(),
				broadcastViewerUrl: ref(''),
			},
			slideShow: { showSetUpSlideShow: closed() },
			presentationProperties: {},
			customShows: [],
			slideCount: 1,
			password: { showPasswordDialog: closed(), isPasswordProtected: closed() },
			fontEmbedding: {
				showFontEmbedding: closed(),
				embedFontsEnabled: closed(),
				usedFontFamilies: ref([]),
				embeddedFontNames: ref([]),
				fontEmbedding: ref({ interactive: false }),
			},
			insertDialogs,
			signatureWorkflow: { showSignatureStripped: closed() },
			signatureCount: 0,
			deckActions: { pendingSlideSizeRescale: ref(null) },
		} as never,
		global: {
			stubs: {
				BroadcastDialog: true,
				SetUpSlideShowDialog: true,
				PasswordProtectionDialog: true,
				FontEmbeddingPanel: true,
				EquationEditorDialog: true,
				SignatureStrippedDialog: true,
				SlideSizeRescalePrompt: true,
			},
		},
	});
	return { wrapper, insertDialogs, onInsertSmartArt };
}

describe('viewer deck dialogs: Insert SmartArt', () => {
	it('opens the shared dialog and routes its choice and close to the viewer', async () => {
		const { wrapper, insertDialogs, onInsertSmartArt } = mountDialogs();
		const dialog = document.body.querySelector('pptx-ui-insert-smartart-dialog') as HTMLElement & {
			open: boolean;
		};
		expect(dialog.open).toBe(false);
		insertDialogs.showInsertSmartArt.value = true;
		await wrapper.vm.$nextTick();
		expect(dialog.open).toBe(true);
		dialog.dispatchEvent(
			new CustomEvent('insert-smartart', { detail: { layout: 'basicCycle', items: ['A', 'B'] } }),
		);
		expect(onInsertSmartArt).toHaveBeenCalledWith('basicCycle', ['A', 'B']);
		dialog.dispatchEvent(new CustomEvent('dialog-close'));
		expect(insertDialogs.showInsertSmartArt.value).toBe(false);
		wrapper.unmount();
	});
});
