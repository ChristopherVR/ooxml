import { mount } from 'svelte';

import { externalSessionRequested } from '../../shared/host-owned-collaboration';
import { oneDriveDemoRequested } from '../../shared/onedrive-demo';
import { sameBrowserRoom } from '../../shared/same-browser-collaboration';
import App from './App.svelte';
import HostOwnedDemo from './HostOwnedDemo.svelte';
import OneDriveDemo from './OneDriveDemo.svelte';
import SameBrowserDemo from './SameBrowserDemo.svelte';

import './styles.css';
import '../../shared/dropzone.css';

/** Demo entry point: mounts the Svelte demo shell. */
const appRoot = document.getElementById('app');
if (!appRoot) {
	throw new Error('missing #app root');
}

mount(
	oneDriveDemoRequested()
		? OneDriveDemo
		: externalSessionRequested()
			? HostOwnedDemo
			: sameBrowserRoom()
				? SameBrowserDemo
				: App,
	{
		target: appRoot,
	},
);
