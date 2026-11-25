import { dirname } from 'node:path';
import type { Options } from 'tsup';
import { commandStrings } from '../xlsx/ui/commands/i18n/index';
import type { CommandLocale } from '../xlsx/ui/commands/i18n/types';

type EsbuildPlugin = NonNullable<Options['esbuildPlugins']>[number];
const localePath = /[\\/]xlsx[\\/]ui[\\/]locales[\\/](en|fr|de|es|zh-CN)[\\/]commands\.ts$/;

function isCommandLocale(value: string | undefined): value is CommandLocale {
	return value === 'en' || value === 'fr' || value === 'de' || value === 'es' || value === 'zh-CN';
}

/** Project the source rows at build time so no browser chunk contains multiple languages. */
export const xlsxCommandLocales: EsbuildPlugin = {
	name: 'xlsx-command-locales',
	setup(build) {
		build.onLoad({ filter: localePath }, ({ path }) => {
			const locale = localePath.exec(path)?.[1];
			if (!isCommandLocale(locale)) return;
			const entries = commandStrings(locale);
			// Keep the English keys once. The flattened tables share their insertion order.
			const contents =
				locale === 'en'
					? `export const commands = Object.fromEntries(${JSON.stringify(Object.keys(entries))}.map(key => [key, key]));`
					: `import { commands as english } from '../en/commands';
const values = ${JSON.stringify(Object.values(entries))};
export const commands = Object.fromEntries(Object.keys(english).map((key, index) => [key, values[index]]));`;
			return { contents, loader: 'js', resolveDir: dirname(path) };
		});
	},
};

/** CJS has no code splitting in esbuild. Load each separately built locale by its output path. */
export const xlsxCjsLocaleImports: EsbuildPlugin = {
	name: 'xlsx-cjs-locale-imports',
	setup(build) {
		build.onResolve(
			{ filter: /^\.\/locales\/(fr|de|es|zh-CN)\/index$/ },
			({ path, importer, kind }) => {
				if (kind === 'dynamic-import' && /[\\/]xlsx[\\/]ui[\\/]localization\.ts$/.test(importer))
					return { path: path.replace(/\/([^/]+)\/index$/, '/xlsx-locale-$1.cjs'), external: true };
			},
		);
	},
};
