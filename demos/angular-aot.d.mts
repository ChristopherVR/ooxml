import type { Plugin } from 'vite';

/** The Angular compiler plugin for an Angular demo, compiling ahead of time (see angular-aot.mjs). */
export function angularAot(packageJson: string, tsconfig: string): Promise<Plugin[]>;
