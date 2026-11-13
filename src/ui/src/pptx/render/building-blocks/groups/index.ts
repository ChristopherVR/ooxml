/**
 * The built-in action groups and the one call that registers all of them on a session.
 *
 * @module render/building-blocks/groups
 */
import type { PptxBuildingBlockSession } from '../session';
import {
	digitalSignaturesGroup,
	documentPropertiesGroup,
	fontEmbeddingGroup,
	passwordProtectionGroup,
} from './file-dialog-groups';
import type {
	DigitalSignaturesPort,
	DocumentPropertiesPort,
	FontEmbeddingPort,
	PasswordProtectionPort,
} from './file-dialog-groups';
import {
	broadcastGroup,
	customShowsGroup,
	setUpSlideShowGroup,
	shortcutsGroup,
} from './show-dialog-groups';
import type { CustomShowsPort, SetUpSlideShowPort, ShortcutsPort } from './show-dialog-groups';

export * from './file-dialog-groups';
export * from './show-dialog-groups';

/** The port of every built-in group, as a binding supplies them (each read at call time). */
export interface BuildingBlockGroupPorts {
	passwordProtection: PasswordProtectionPort;
	documentProperties: DocumentPropertiesPort;
	fontEmbedding: FontEmbeddingPort;
	digitalSignatures: DigitalSignaturesPort;
	setUpSlideShow: SetUpSlideShowPort;
	customShows: CustomShowsPort;
	shortcuts: ShortcutsPort;
}

/** The `actions` keys of the built-in groups, which every binding's `actions` carries. */
export const BUILDING_BLOCK_GROUP_IDS = [
	'passwordProtection',
	'documentProperties',
	'fontEmbedding',
	'digitalSignatures',
	'setUpSlideShow',
	'customShows',
	'broadcast',
	'shortcuts',
] as const;

/** Register every built-in group on `session`; returns the function that unregisters them. */
export function registerBuildingBlockGroups(
	session: Pick<PptxBuildingBlockSession, 'register'>,
	ports: BuildingBlockGroupPorts,
): () => void {
	const stops = [
		session.register(passwordProtectionGroup, ports.passwordProtection),
		session.register(documentPropertiesGroup, ports.documentProperties),
		session.register(fontEmbeddingGroup, ports.fontEmbedding),
		session.register(digitalSignaturesGroup, ports.digitalSignatures),
		session.register(setUpSlideShowGroup, ports.setUpSlideShow),
		session.register(customShowsGroup, ports.customShows),
		session.register(broadcastGroup, null),
		session.register(shortcutsGroup, ports.shortcuts),
	];
	return () => {
		for (const stop of stops) {
			stop();
		}
	};
}
