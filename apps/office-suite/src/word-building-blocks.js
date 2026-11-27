import { createBuildingBlocksStore } from 'ooxml-ui/suite';

const profiles = new Map();

/** Word's shared building blocks for a profile, kept on this device (see PRIVACY.md). */
export function wordBuildingBlocks(profile) {
	if (!profiles.has(profile)) {
		profiles.set(profile, createBuildingBlocksStore(`ooxml-suite-building-blocks-${profile}`));
	}
	return profiles.get(profile);
}
