import { mount } from 'svelte';
import App from './App.svelte';
import { createWorkspace } from '../demo/workspace';

mount(App, {
	target: window.document.getElementById('viewer')!,
	props: { workspace: createWorkspace() },
});
