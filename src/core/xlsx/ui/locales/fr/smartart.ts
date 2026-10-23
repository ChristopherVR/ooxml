// The dotted keys of the shared SmartArt UI strings (`ooxml-core/diagram`): the Insert SmartArt
// catalogue's category, description and layout names. The plain keys are in
// ../../commands/i18n/smartart.ts with the other command strings.
import { DIAGRAM_UI_STRINGS } from '../../../../diagram/strings/index';

export const smartart: Record<string, string> = Object.fromEntries(
	Object.entries(DIAGRAM_UI_STRINGS.fr).filter(([key]) => key.startsWith('smartart.')),
);
