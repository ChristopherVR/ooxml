import { createApp, h, ref } from 'vue';
import { Teams } from 'openteams-vue-viewer';
import { config, showStaticNotice, userId, userName, workspaceId } from '../../shared';
import { currentHostClass, onHostClass, recordOpenFile } from '../../test-hooks';

// `class` falls through to <teams-app>; the browser tests swap it (../../test-hooks).
const hostClass = ref(currentHostClass());
onHostClass((value) => (hostClass.value = value));

showStaticNotice();
createApp({
	render: () =>
		h(Teams, {
			class: hostClass.value,
			workspaceId,
			userName,
			userId,
			config,
			onOpenFile: recordOpenFile,
		}),
}).mount('#app');
