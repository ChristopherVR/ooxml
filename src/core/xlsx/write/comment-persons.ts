import { NS, buildXml, first, parseXml } from '../../xml/index';
import { guidFrom } from './comment-ids';
import { XML_HEADER } from './xml-out';

/** Retain source account identities, including people who share a display name. */
export class PersonRegistry {
	private readonly names = new Map<string, string>();
	private readonly byName = new Map<string, string>();
	private readonly identities = new Map<string, string>();
	private readonly added = new Map<string, string>();
	private readonly retained = new Map<string, string>();
	constructor(
		existing?: ReadonlyMap<string, string>,
		private readonly source?: string,
	) {
		for (const [id, name] of existing ?? []) {
			this.names.set(id, name);
			this.identities.set(id.toUpperCase(), id);
			if (!this.byName.has(name)) this.byName.set(name, id);
		}
	}
	get changed(): boolean {
		return this.added.size > 0;
	}
	/** Bring referenced account records along with a thread copied from another workbook. */
	retain(records: Readonly<Record<string, string>> | undefined): void {
		for (const [sourceId, xml] of Object.entries(records ?? {})) {
			const known = this.identities.get(sourceId.toUpperCase());
			if (known && (this.source || this.retained.has(known))) continue;
			const node = parseXml(xml).documentElement;
			if (
				node.localName !== 'person' ||
				node.namespaceURI !== NS.tc ||
				node.getAttribute('id')?.toUpperCase() !== sourceId.toUpperCase()
			)
				continue;
			const name = node.getAttribute('displayName') ?? '';
			const id = known ?? node.getAttribute('id')!;
			this.retained.set(id, xml);
			this.names.set(id, name);
			this.identities.set(id.toUpperCase(), id);
			this.added.set(id, name);
			if (!this.byName.has(name)) this.byName.set(name, id);
		}
	}
	id(name: string, preferred?: string): string {
		const preferredId = preferred ? this.identities.get(preferred.toUpperCase()) : undefined;
		if (preferredId && this.names.get(preferredId) === name) return preferredId;
		let id = this.byName.get(name);
		if (!id) {
			let index = 0;
			id = guidFrom(`person:${name}`);
			while (this.identities.has(id.toUpperCase())) id = guidFrom(`person:${name}:${++index}`);
			this.byName.set(name, id);
			this.names.set(id, name);
			this.identities.set(id.toUpperCase(), id);
			this.added.set(id, name);
		}
		return id;
	}
	xml(): string {
		if (this.source && !this.changed) return this.source;
		const doc = parseXml(
			this.source ?? `${XML_HEADER}<personList xmlns="${NS.tc}" xmlns:x="${NS.x}"/>`,
		);
		const root = doc.documentElement;
		for (const [id, name] of this.source ? this.added : this.names) {
			const retained = this.retained.get(id);
			if (retained) {
				root.insertBefore(
					doc.importNode(parseXml(retained).documentElement, true),
					first(root, 'extLst', NS.tc) ?? null,
				);
				continue;
			}
			const person = doc.createElementNS(NS.tc, root.prefix ? `${root.prefix}:person` : 'person');
			for (const [key, value] of Object.entries({
				id,
				displayName: name,
				userId: name,
				providerId: 'None',
			})) {
				person.setAttribute(key, value);
			}
			root.insertBefore(person, first(root, 'extLst', NS.tc) ?? null);
		}
		return buildXml(doc);
	}
}
