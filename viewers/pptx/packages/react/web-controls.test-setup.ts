import { getI18n, setI18n } from 'react-i18next';
import { afterAll } from 'vitest';

// Import the dedicated module so tests mocking the shared barrel or core keep
// their own module setup. This file lives outside src and is test-only.
import { registerPptxWebControls } from '../../../../src/ui/src/pptx/web-components';

registerPptxWebControls();

// initReactI18next writes a process-wide fallback even when a test uses its own
// provider. Restore it when this file ends so the next shared-worker file does
// not inherit the last locale tested here.
const previousI18n = getI18n();
afterAll(() => setI18n(previousI18n));
