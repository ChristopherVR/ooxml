// Plain-JavaScript half of the Svelte package (`openteams-svelte-viewer/runtime`): the raw store
// plus the helpers `Teams.svelte` imports. The build bundles the web component into it, and the
// published `Teams.svelte` imports from `./runtime.js` instead of `teams-viewer`.
export * from 'teams-viewer';
export { teamsStore } from './store';
