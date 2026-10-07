import { runBuild } from './run-build.mjs';
import { rewriteEsmDeclarationImports } from '../../../scripts/esm-declarations.mjs';

// Use the strict declaration compiler, as the full build does. tsup's legacy
// declaration plugin cannot use the TypeScript 7 compiler API.
await Promise.all([
	runBuild('tsc', ['-p', 'tsconfig.build.json']),
	runBuild('tsup', ['--config', 'tsup.visio.config.ts']),
]);
rewriteEsmDeclarationImports('dist');
