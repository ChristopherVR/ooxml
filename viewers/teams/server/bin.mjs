#!/usr/bin/env node
// `openteams-server`: the command installed by the package. A separate file because an npm bin is
// started through a link, so `index.mjs` cannot tell it is the entry point from `process.argv[1]`.
import { runTeamsServer } from './index.mjs';

await runTeamsServer();
