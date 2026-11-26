import { completeDemoSignInPopup } from 'ooxml-demo-onedrive';

import '../../shared/zod-jitless';
import { mount } from 'svelte';

// Test-only: `?elementRenderFault=<id>` makes that element's renderer throw (e2e).
import '../../shared/element-render-fault';
import { enableDemoAddInPreferences } from '../../shared/demo-add-ins';
import { externalSessionRequested } from '../../shared/host-owned-collaboration';
import { sameBrowserRoom } from '../../shared/same-browser-collaboration';
import App from './App.svelte';
import BuildingBlocksDemo from './BuildingBlocksDemo.svelte';
import HostOwnedDemo from './HostOwnedDemo.svelte';
import SameBrowserDemo from './SameBrowserDemo.svelte';

import './styles.css';
import '../../shared/dropzone.css';

/** Demo entry point: mounts the Svelte demo shell. */
const appRoot = document.getElementById('app');
if (!appRoot) {
	throw new Error('missing #app root');
}

if (!completeDemoSignInPopup()) {
	enableDemoAddInPreferences(appRoot);
	mount(
		new URLSearchParams(window.location.search).get('blocks') === '1'
			? BuildingBlocksDemo
			: externalSessionRequested()
				? HostOwnedDemo
				: sameBrowserRoom()
					? SameBrowserDemo
					: App,
		{
			target: appRoot,
		},
	);
}
