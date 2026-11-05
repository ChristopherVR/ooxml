import { mount } from 'svelte';
import App from './App.svelte';
import { completeDemoSignInPopup } from 'ooxml-demo-onedrive';
import { createWorkspace } from '../demo/workspace';

// In the MSAL sign-in popup the page only hands the response back to the opener.
if (!completeDemoSignInPopup()) {
	mount(App, {
		target: window.document.getElementById('viewer')!,
		props: { workspace: createWorkspace() },
	});
}
