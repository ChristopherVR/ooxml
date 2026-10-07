import { createApp, defineComponent, h, onMounted, shallowRef } from 'vue';
import { VisioViewer } from '../../../viewers/visio/packages/bindings/src/vue';
import type { ViewerHandle } from '../../../viewers/visio/packages/bindings/src/common';
import { createWorkspace } from '../demo/workspace';

// The Vue demo: the VisioViewer component with its exposed handle and the document as a ref.
const workspace = createWorkspace();

const App = defineComponent({
	setup() {
		const viewer = shallowRef<ViewerHandle | null>(null);
		const document = shallowRef(workspace.initialDocument);
		onMounted(() => {
			if (viewer.value) workspace.attach(viewer.value, (next) => (document.value = next));
		});
		return () =>
			h(VisioViewer, { ref: viewer, document: document.value, events: workspace.events });
	},
});

createApp(App).mount(window.document.getElementById('viewer')!);
