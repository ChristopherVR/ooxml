import { createApp, h } from 'vue';
import { Teams } from 'openteams-vue-viewer';
import { config, showStaticNotice, userId, userName, workspaceId } from '../../shared';

showStaticNotice();
createApp({
	render: () => h(Teams, { workspaceId, userName, userId, config }),
}).mount('#app');
