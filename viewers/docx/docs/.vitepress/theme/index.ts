import type { Theme } from 'vitepress';
import DefaultTheme from 'vitepress/theme';

import Layout from './Layout.vue';
import LandingHome from './landing/LandingHome.vue';

// Self-hosted (bundled into the site), so no page view contacts a font CDN.
import '@fontsource-variable/bricolage-grotesque/opsz.css';
import '@fontsource/ibm-plex-mono/400.css';
import '@fontsource/ibm-plex-mono/400-italic.css';
import '@fontsource/ibm-plex-mono/500.css';

import './custom.css';
import './landing/landing.css';

export default {
	extends: DefaultTheme,
	Layout,
	enhanceApp({ app }) {
		app.component('LandingHome', LandingHome);
	},
} satisfies Theme;
