/** The document stencil: the masters a drawing defines, and the parts that hold them. */
import type { Size } from './geometry';
import { type ShapeDef, drawnShape, shapeCount } from './shape';
import type { StyleBook } from './styles';
import { REL_NS, VISIO_NS, XML_HEADER, attrs, cell } from './xml';

export interface MasterDef {
	name: string;
	/** The tip Visio shows for the master in the Shapes window. */
	prompt: string;
	size: Size;
	shape: ShapeDef;
	/** Text the master itself carries (its instances inherit it unless they have their own). */
	text?: string;
	/** A complete master shape written by its owner (1-D connectors); replaces `shape`. */
	custom?: (styles: StyleBook, id: number) => string;
}

/** The first shape ID inside a master part; sub-shapes follow depth first. */
export const MASTER_ROOT_ID = 5;

/**
 * A stable GUID in the form Visio keeps (`{8-4-4-4-12}`), derived from a name: the samples must
 * not depend on the runtime's random UUIDs.
 */
export function stableGuid(seed: string): string {
	let digits = '';
	for (let round = 0; digits.length < 32; round++) {
		// FNV-1a, 32 bits, salted per round.
		let hash = 0x811c9dc5 ^ round;
		for (const char of `${round}:${seed}`) {
			hash ^= char.codePointAt(0) ?? 0;
			hash = Math.imul(hash, 0x01000193) >>> 0;
		}
		digits += hash.toString(16).padStart(8, '0');
	}
	const upper = digits.toUpperCase();
	return `{${upper.slice(0, 8)}-${upper.slice(8, 12)}-4${upper.slice(13, 16)}-8${upper.slice(17, 20)}-${upper.slice(20, 32)}}`;
}

/** A master with the ID it has in the drawing. */
export interface PlacedMaster extends MasterDef {
	id: number;
	/** How many instances the pages hold so far, for Visio's `Name.ID` naming. */
	uses: number;
}

export class Stencil {
	readonly masters: PlacedMaster[];
	private readonly byName = new Map<string, PlacedMaster>();

	constructor(defs: readonly MasterDef[]) {
		this.masters = defs.map((def, index) => ({ ...def, id: index + 1, uses: 0 }));
		for (const master of this.masters) {
			if (this.byName.has(master.name)) throw new Error(`Master ${master.name} is defined twice.`);
			this.byName.set(master.name, master);
		}
	}

	get(name: string): PlacedMaster {
		const master = this.byName.get(name);
		if (!master) throw new Error(`Unknown master: ${name}`);
		return master;
	}

	/** `visio/masters/masters.xml`. */
	mastersXml(themeCells = ''): string {
		const entries = this.masters.map((master) => {
			const sheet =
				'<PageSheet LineStyle="0" FillStyle="0" TextStyle="0">' +
				themeCells +
				cell('PageWidth', master.size.w) +
				cell('PageHeight', master.size.h) +
				cell('PageScale', 1) +
				cell('DrawingScale', 1) +
				cell('DrawingSizeType', 4) +
				cell('DrawingScaleType', 0) +
				'</PageSheet>';
			const head = attrs({
				ID: master.id,
				NameU: master.name,
				IsCustomNameU: 1,
				Name: master.name,
				IsCustomName: 1,
				Prompt: master.prompt,
				IconSize: 1,
				AlignName: 2,
				MatchByName: 0,
				IconUpdate: 1,
				UniqueID: stableGuid(`master:${master.name}`),
				BaseID: stableGuid(`base:${master.name}`),
				PatternFlags: 0,
				Hidden: 0,
				MasterType: 2,
			});
			return `<Master${head}>${sheet}<Rel r:id="rId${master.id}"/></Master>`;
		});
		return `${XML_HEADER}<Masters xmlns="${VISIO_NS}" xmlns:r="${REL_NS}" xml:space="preserve">${entries.join('')}</Masters>`;
	}

	/** `visio/masters/_rels/masters.xml.rels`. */
	relsXml(): string {
		const type = 'http://schemas.microsoft.com/visio/2010/relationships/master';
		const entries = this.masters.map(
			(master) =>
				`<Relationship Id="rId${master.id}" Type="${type}" Target="master${master.id}.xml"/>`,
		);
		return `${XML_HEADER}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${entries.join('')}</Relationships>`;
	}

	/** One `visio/masters/master<N>.xml`. */
	masterXml(master: PlacedMaster, styles: StyleBook): string {
		let next = MASTER_ROOT_ID;
		const shape = master.custom
			? master.custom(styles, MASTER_ROOT_ID)
			: drawnShape(
					master.shape,
					{ x: master.size.w / 2, y: master.size.h / 2, size: master.size },
					styles,
					() => next++,
					{ body: master.text ? `<Text>${master.text}</Text>` : '' },
				);
		return `${XML_HEADER}<MasterContents xmlns="${VISIO_NS}" xmlns:r="${REL_NS}" xml:space="preserve"><Shapes>${shape}</Shapes></MasterContents>`;
	}
}

/** How many shape IDs an instance of the master takes on a page. */
export const instanceIdCount = (master: MasterDef): number =>
	master.custom ? 1 : shapeCount(master.shape);
