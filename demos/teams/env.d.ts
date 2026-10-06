// The one build-time flag the demos read. `scripts/build-pages.mjs` sets VITE_TEAMS_STATIC=1 for the
// GitHub Pages build, where no server exists: the demos then start in local (same-browser) mode and
// say so on the page.
interface ImportMetaEnv {
	readonly VITE_TEAMS_STATIC?: string;
}

interface ImportMeta {
	readonly env: ImportMetaEnv;
}

declare module '*.css';
