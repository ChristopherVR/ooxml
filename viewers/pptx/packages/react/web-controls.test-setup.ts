import { getI18n, setI18n } from 'react-i18next';
import { afterAll, beforeAll } from 'vitest';

// Import the dedicated module so tests mocking the shared barrel or core keep
// their own module setup. This file lives outside src and is test-only.
import { registerPptxWebControls } from '../../../../src/ui/src/pptx/web-components';

registerPptxWebControls();

// The main entries register react-i18next as the fallback translation source when they load;
// specs import source modules directly, so register it for every file. In `beforeAll`, after
// the file's imports: a file that resets the module registry first registers in its own graph.
beforeAll(async () => {
	await import('./src/viewer/i18n/react-i18next-source');
});

// initReactI18next writes a process-wide fallback even when a test uses its own
// provider. Restore it when this file ends so the next shared-worker file does
// not inherit the last locale tested here.
const previousI18n = getI18n();
afterAll(() => setI18n(previousI18n));
