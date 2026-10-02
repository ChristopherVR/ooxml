import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

export const viewerRoot = resolve(dirname(fileURLToPath(import.meta.url)), '..');
// Keep the portable file:../ooxml dependency. An override is local to this process;
// use link:core after npm ci to point the installed dependency at that checkout.
export const coreDirectory = resolve(viewerRoot, process.env.VISIO_CORE_DIR ?? '../ooxml');
