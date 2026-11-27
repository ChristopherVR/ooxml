/**
 * A hand-written `p:timing` tree for the slides that build their text one paragraph at a time.
 * The core's animation writer records a by-paragraph sequence as one whole-shape effect plus a
 * `p:bldP`, which PowerPoint plays (and reports) as a single effect; a real paragraph build needs
 * one effect per paragraph, each targeting `p:txEl/p:pRg`. Slides that need that use this writer.
 */

export type FxKind = 'appear' | 'fade' | 'rise' | 'fly' | 'wipe' | 'fadeOut';

export interface Fx {
	/** The shape's name on the slide; resolved to its id after the core has saved the slide. */
	target: string;
	/** 0-based paragraph of the shape's text; the whole shape when omitted. */
	paragraph?: number;
	kind: FxKind;
	/** `click` waits for a click, `with` starts with the previous effect, `after` when it ends. */
	start?: 'click' | 'with' | 'after';
	/** Milliseconds. */
	duration?: number;
	delay?: number;
}

const PRESETS: Record<FxKind, { id: number; subtype: number; exit?: boolean }> = {
	appear: { id: 1, subtype: 0 },
	fade: { id: 10, subtype: 0 },
	rise: { id: 42, subtype: 0 },
	fly: { id: 2, subtype: 4 },
	wipe: { id: 22, subtype: 8 },
	fadeOut: { id: 10, subtype: 0, exit: true },
};

class Ids {
	private next = 1;
	take(): number {
		return this.next++;
	}
}

function behaviours(kind: FxKind, target: string, duration: number, ids: Ids): string {
	const common = (dur: number | string, extra = '') =>
		`<p:cBhvr${extra}><p:cTn id="${ids.take()}" dur="${dur}" fill="hold"/><p:tgtEl>${target}</p:tgtEl>`;
	const show = (value: string, delay: number) =>
		`<p:set><p:cBhvr><p:cTn id="${ids.take()}" dur="1" fill="hold"><p:stCondLst><p:cond delay="${delay}"/></p:stCondLst></p:cTn><p:tgtEl>${target}</p:tgtEl><p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="${value}"/></p:to></p:set>`;
	const effect = (filter: string, direction: 'in' | 'out') =>
		`<p:animEffect transition="${direction}" filter="${filter}"><p:cBhvr><p:cTn id="${ids.take()}" dur="${duration}"/><p:tgtEl>${target}</p:tgtEl></p:cBhvr></p:animEffect>`;
	const move = (attribute: string, from: string, to: string) =>
		`<p:anim calcmode="lin" valueType="num">${common(duration, ' additive="base"')}<p:attrNameLst><p:attrName>${attribute}</p:attrName></p:attrNameLst></p:cBhvr><p:tavLst><p:tav tm="0"><p:val><p:strVal val="${from}"/></p:val></p:tav><p:tav tm="100000"><p:val><p:strVal val="${to}"/></p:val></p:tav></p:tavLst></p:anim>`;
	switch (kind) {
		case 'appear':
			return show('visible', 0);
		case 'fade':
			return show('visible', 0) + effect('fade', 'in');
		case 'wipe':
			return show('visible', 0) + effect('wipe(left)', 'in');
		case 'rise':
			return (
				show('visible', 0) +
				effect('fade', 'in') +
				move('ppt_x', '#ppt_x', '#ppt_x') +
				move('ppt_y', '#ppt_y+.1', '#ppt_y')
			);
		case 'fly':
			return (
				show('visible', 0) +
				move('ppt_x', '#ppt_x', '#ppt_x') +
				move('ppt_y', '1+#ppt_h/2', '#ppt_y')
			);
		case 'fadeOut':
			return effect('fade', 'out') + show('hidden', Math.max(0, duration - 1));
	}
}

/**
 * The `p:timing` element for a main sequence of effects. `shapeIds` maps a shape name to the
 * `p:cNvPr/@id` the slide was saved with.
 */
export function timingXml(effects: readonly Fx[], shapeIds: ReadonlyMap<string, string>): string {
	const ids = new Ids();
	const root = ids.take();
	const main = ids.take();
	// Click groups, each a list of "with" groups that start one after the other.
	const clicks: { auto: boolean; groups: { delay: number; xml: string[]; length: number }[] }[] =
		[];
	const built = new Set<string>();
	for (const [index, fx] of effects.entries()) {
		const spid = shapeIds.get(fx.target);
		if (!spid) throw new Error(`Animation target "${fx.target}" is not a shape on its slide.`);
		const start = fx.start ?? 'click';
		const duration = fx.duration ?? 500;
		const delay = fx.delay ?? 0;
		if (start === 'click' || index === 0) clicks.push({ auto: start !== 'click', groups: [] });
		const click = clicks[clicks.length - 1]!;
		const last = click.groups[click.groups.length - 1];
		if (!last || start === 'after')
			click.groups.push({ delay: last ? last.delay + last.length : 0, xml: [], length: 0 });
		const group = click.groups[click.groups.length - 1]!;
		group.length = Math.max(group.length, delay + duration);
		const target =
			fx.paragraph === undefined
				? `<p:spTgt spid="${spid}"/>`
				: `<p:spTgt spid="${spid}"><p:txEl><p:pRg st="${fx.paragraph}" end="${fx.paragraph}"/></p:txEl></p:spTgt>`;
		if (fx.paragraph !== undefined) built.add(spid);
		const preset = PRESETS[fx.kind];
		const node =
			start === 'click' ? 'clickEffect' : start === 'with' ? 'withEffect' : 'afterEffect';
		const effectId = ids.take();
		group.xml.push(
			`<p:par><p:cTn id="${effectId}" presetID="${preset.id}" presetClass="${preset.exit ? 'exit' : 'entr'}" presetSubtype="${preset.subtype}" fill="hold" grpId="0" nodeType="${node}">` +
				`<p:stCondLst><p:cond delay="${delay}"/></p:stCondLst><p:childTnLst>${behaviours(fx.kind, target, duration, ids)}</p:childTnLst></p:cTn></p:par>`,
		);
	}
	const sequence = clicks
		.map((click) => {
			const begin = click.auto
				? `<p:cond delay="indefinite"/><p:cond evt="onBegin" delay="0"><p:tn val="${main}"/></p:cond>`
				: '<p:cond delay="indefinite"/>';
			const groups = click.groups
				.map(
					(group) =>
						`<p:par><p:cTn id="${ids.take()}" fill="hold"><p:stCondLst><p:cond delay="${group.delay}"/></p:stCondLst><p:childTnLst>${group.xml.join('')}</p:childTnLst></p:cTn></p:par>`,
				)
				.join('');
			return `<p:par><p:cTn id="${ids.take()}" fill="hold"><p:stCondLst>${begin}</p:stCondLst><p:childTnLst>${groups}</p:childTnLst></p:cTn></p:par>`;
		})
		.join('');
	const builds = [...built].map((spid) => `<p:bldP spid="${spid}" grpId="0" build="p"/>`).join('');
	return (
		`<p:timing><p:tnLst><p:par><p:cTn id="${root}" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst>` +
		`<p:seq concurrent="1" nextAc="seek"><p:cTn id="${main}" dur="indefinite" nodeType="mainSeq"><p:childTnLst>${sequence}</p:childTnLst></p:cTn>` +
		'<p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst>' +
		'<p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst></p:seq>' +
		`</p:childTnLst></p:cTn></p:par></p:tnLst>${builds ? `<p:bldLst>${builds}</p:bldLst>` : ''}</p:timing>`
	);
}
