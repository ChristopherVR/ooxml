import { completeDemoSignInPopup } from 'ooxml-demo-onedrive';

import '../../shared/zod-jitless';
import { createApp } from 'vue';

// Test-only: `?elementRenderFault=<id>` makes that element's renderer throw (e2e).
import '../../shared/element-render-fault';
import { enableDemoAddInPreferences } from '../../shared/demo-add-ins';
import { externalSessionRequested } from '../../shared/host-owned-collaboration';
import { sameBrowserRoom } from '../../shared/same-browser-collaboration';
import App from './App.vue';
import BuildingBlocksDemo from './BuildingBlocksDemo.vue';
import HostOwnedDemo from './HostOwnedDemo.vue';
import i18n from './i18n';
import SameBrowserDemo from './SameBrowserDemo.vue';

import '../../shared/dropzone.css';

if (!completeDemoSignInPopup()) {
	const root = document.getElementById('app');
	if (root) enableDemoAddInPreferences(root);
	createApp(
		new URLSearchParams(window.location.search).get('blocks') === '1'
			? BuildingBlocksDemo
			: externalSessionRequested()
				? HostOwnedDemo
				: sameBrowserRoom()
					? SameBrowserDemo
					: App,
	)
		.use(i18n)
		.mount('#app');
}
