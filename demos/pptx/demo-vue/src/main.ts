import { createApp } from 'vue';

import { externalSessionRequested } from '../../shared/host-owned-collaboration';
import { sameBrowserRoom } from '../../shared/same-browser-collaboration';
import App from './App.vue';
import HostOwnedDemo from './HostOwnedDemo.vue';
import i18n from './i18n';
import SameBrowserDemo from './SameBrowserDemo.vue';

import '../../shared/dropzone.css';

createApp(externalSessionRequested() ? HostOwnedDemo : sameBrowserRoom() ? SameBrowserDemo : App)
	.use(i18n)
	.mount('#app');
