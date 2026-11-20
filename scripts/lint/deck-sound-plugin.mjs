/**
 * Oxlint plugin: a PowerPoint viewer plays a deck's sounds only through
 * `playDeckSound` (ooxml-ui/pptx `render/deck-sound.ts`), which withholds a
 * linked sound until the viewer allows linked content. `no-restricted-globals`
 * catches a bare `Audio`; this rule catches the other ways to make an audio
 * element: `Audio` read off any object (`window.Audio`, `globalThis.Audio`,
 * `self['Audio']`) and `createElement('audio')` / `createElementNS(ns, 'audio')`
 * on any document.
 */

/** The string a literal or expression-free template literal holds, if any. */
function staticString(node) {
	if (!node) {
		return undefined;
	}
	if (node.type === 'Literal' && typeof node.value === 'string') {
		return node.value;
	}
	if (node.type === 'TemplateLiteral' && node.expressions.length === 0) {
		return node.quasis[0]?.value.cooked ?? undefined;
	}
	return undefined;
}

function propertyName(member) {
	return member.computed ? staticString(member.property) : member.property.name;
}

const noUnsanctionedAudio = {
	meta: {
		type: 'problem',
		docs: { description: 'Create audio elements only in the deck-sound sink.' },
		messages: {
			audio:
				"Play deck sounds with playDeckSound (ooxml-ui/pptx render/deck-sound.ts): it applies the viewer's linked-content consent.",
		},
		schema: [],
	},
	create(context) {
		return {
			MemberExpression(node) {
				if (propertyName(node) === 'Audio') {
					context.report({ node, messageId: 'audio' });
				}
			},
			CallExpression(node) {
				const callee = node.callee;
				if (callee.type !== 'MemberExpression') {
					return;
				}
				const name = propertyName(callee);
				const tag =
					name === 'createElement'
						? staticString(node.arguments[0])
						: name === 'createElementNS'
							? staticString(node.arguments[1])
							: undefined;
				if (tag?.toLowerCase() === 'audio') {
					context.report({ node, messageId: 'audio' });
				}
			},
		};
	},
};

export default {
	meta: { name: 'deck-sound' },
	rules: { 'no-unsanctioned-audio': noUnsanctionedAudio },
};
