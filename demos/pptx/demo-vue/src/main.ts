import { createApp } from 'vue';

import { externalSessionRequested } from '../../shared/host-owned-collaboration';
import { oneDriveDemoRequested } from '../../shared/onedrive-demo';
import { sameBrowserRoom } from '../../shared/same-browser-collaboration';
import App from './App.vue';
import HostOwnedDemo from './HostOwnedDemo.vue';
import i18n from './i18n';
import OneDriveDemo from './OneDriveDemo.vue';
import SameBrowserDemo from './SameBrowserDemo.vue';

import '../../shared/dropzone.css';

createApp(
	oneDriveDemoRequested()
		? OneDriveDemo
		: externalSessionRequested()
			? HostOwnedDemo
			: sameBrowserRoom()
				? SameBrowserDemo
				: App,
)
	.use(i18n)
	.mount('#app');
