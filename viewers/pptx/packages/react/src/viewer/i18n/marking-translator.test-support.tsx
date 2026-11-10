/**
 * marking-translator.test-support.tsx: a React root that renders every tree under a
 * `PptxI18nProvider` whose translator marks each key (`t:<key>`). A spec then asserts the
 * marked text, which proves that the component translates the right key through `t`: one that
 * renders hardcoded English or reads the English dictionary directly shows unmarked text and
 * fails. The provider also keeps the spec independent of react-i18next's global instance.
 *
 * Not a `.test.tsx` file on purpose: a shared helper is not a suite.
 */
import type { ReactNode } from 'react';
import { createRoot } from 'react-dom/client';
import type { Root } from 'react-dom/client';

import { PptxI18nProvider } from './provider';

/**
 * The text the marking translator returns for a key. It differs from the key, since the provider
 * treats a translator that returns the key itself as a miss and falls back to English.
 */
export const marked = (key: string): string => `t:${key}`;

/** `createRoot`, with every `render` wrapped in the marking `PptxI18nProvider`. */
export function createMarkedRoot(container: Element): Root {
	const root = createRoot(container);
	return {
		render: (children: ReactNode) =>
			root.render(
				<PptxI18nProvider translate={marked} locale='en'>
					{children}
				</PptxI18nProvider>,
			),
		unmount: () => root.unmount(),
	};
}
