import { createApp, defineComponent, h, onMounted, shallowRef, type PropType } from 'vue';
import { VisioViewer } from '../../../viewers/visio/packages/bindings/src/vue';
import type { ViewerHandle } from '../../../viewers/visio/packages/bindings/src/common';
import { completeDemoSignInPopup } from 'ooxml-demo-onedrive';
import { createWorkspace, type Workspace } from '../demo/workspace';

// The Vue demo: the VisioViewer component with its exposed handle.
const popup = completeDemoSignInPopup();
// In the MSAL sign-in popup the page only hands the response back to the opener.
const workspace = popup ? undefined : createWorkspace();

const App = defineComponent({
	props: { workspace: { type: Object as PropType<Workspace>, required: true } },
	setup(props) {
		const viewer = shallowRef<ViewerHandle | null>(null);
		onMounted(() => {
			if (viewer.value) props.workspace.attach(viewer.value);
		});
		return () =>
			h(VisioViewer, {
				ref: viewer,
				document: props.workspace.initialDocument,
				fileProviders: props.workspace.fileProviders,
				events: props.workspace.events,
				'aria-label': 'Visio diagram',
			});
	},
});

if (workspace) createApp(App, { workspace }).mount(window.document.getElementById('viewer')!);
