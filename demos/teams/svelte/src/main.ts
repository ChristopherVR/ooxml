import { mount } from 'svelte';
import Teams from 'openteams-svelte-viewer';
import { config, showStaticNotice, userId, userName, workspaceId } from '../../shared';

showStaticNotice();
mount(Teams, {
	target: document.getElementById('app')!,
	props: { workspaceId, userName, userId, config },
});
