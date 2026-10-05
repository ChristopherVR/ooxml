#!/usr/bin/env node
/**
 * npm-trusted-publishers.mjs: print the `npm trust` commands that point every package of this
 * repository's release table at this repository's release workflow.
 *
 * npm trusted publishing (OIDC) is configured per package on npmjs.com: the publisher entry names
 * the repository, the workflow file and the environment the publish runs in. The viewer packages
 * were published from their own repositories, so each still trusts the old repository until it is
 * switched to `ChristopherVR/ooxml`, `release.yml`, environment `npm`.
 *
 *   node scripts/npm-trusted-publishers.mjs
 *
 * This only prints; nothing is changed. Review the list, then run the commands yourself, logged in
 * as a maintainer of every package, with npm >= 12 (`npm trust` is not in npm 11) and a token or
 * session that may change package settings. `npm trust list <package>` shows what a package trusts
 * today, and `npm trust revoke <package> --id=<trust-id>` removes the old entry once the new one
 * works. A package that has never been published cannot have a publisher yet.
 */

import { PACKAGES } from './release-plan.mjs';

const REPOSITORY = 'ChristopherVR/ooxml';
const WORKFLOW = 'release.yml';
const ENVIRONMENT = 'npm';

for (const meta of Object.values(PACKAGES)) {
	console.log(
		`npm trust github ${meta.npm} --repo ${REPOSITORY} --file ${WORKFLOW} --env ${ENVIRONMENT} --yes`,
	);
}
