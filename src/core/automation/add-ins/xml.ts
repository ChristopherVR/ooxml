import { elements, parseXml } from '../../xml/xml';
import type { AddInIssue, AddInManifest, AddInReadOptions } from './model';
import { OFFICE_ADD_IN_HOSTS } from './model';
import {
	AddInManifestError,
	declaredOrigin,
	guid,
	invalid,
	localeTag,
	MANIFEST_DEPTH,
	missingRequirements,
	origin,
	PERMISSIONS,
	text,
	unique,
	url,
	version,
} from './validation';
import {
	all,
	BASE,
	OVERRIDE,
	one,
	typeName,
	value,
	variants,
	XML_BASE_HOSTS,
	xmlRequirements,
} from './xml-fields';
import { readXmlOverrides } from './xml-overrides';

// TaskPaneApp sequence from the public Office Add-in manifest XSD.
const taskPaneOrder = [
	'Id',
	'AlternateId',
	'Version',
	'ProviderName',
	'DefaultLocale',
	'DisplayName',
	'Description',
	'IconUrl',
	'HighResolutionIconUrl',
	'SupportUrl',
	'AppDomains',
	'Hosts',
	'Requirements',
	'DefaultSettings',
	'Permissions',
	'Dictionary',
	'VersionOverrides',
	'ExtendedOverrides',
];

function validateTaskPaneOrder(root: Element): void {
	let previous = -1;
	for (const child of elements(root)) {
		const namespace = child.localName === 'VersionOverrides' ? OVERRIDE : BASE;
		const index = taskPaneOrder.indexOf(child.localName);
		if (index <= previous || child.namespaceURI !== namespace) invalid(child.localName);
		previous = index;
	}
}

export function readXmlManifest(source: string, options: AddInReadOptions): AddInManifest {
	let root: Element;
	try {
		root = parseXml(source, {
			label: 'add-in manifest',
			limits: { maxDepth: MANIFEST_DEPTH, maxNodes: 8192 },
			warningsAsErrors: true,
		}).documentElement;
	} catch {
		return invalid('XML syntax, DTD or nesting limit');
	}
	if (root.localName !== 'OfficeApp') return invalid('OfficeApp');
	if (root.namespaceURI !== BASE) throw new AddInManifestError('version', 'OfficeApp');
	if (typeName(root) !== 'TaskPaneApp') throw new AddInManifestError('extension', 'OfficeApp');
	validateTaskPaneOrder(root);
	const id = guid(value(root, 'Id'), 'Id');
	const release = version(value(root, 'Version'), 'Version');
	const publisher = text(value(root, 'ProviderName'), 'ProviderName', 256);
	const locale = localeTag(value(root, 'DefaultLocale'));
	const localize = (node: Element | undefined, field: string, convert?: (v: unknown) => string) => {
		if (!node) return invalid(field);
		return variants(node, convert)(options.locale, locale);
	};
	const hostsNode = one(root, 'Hosts');
	const names = all(hostsNode, 'Host').map((host) => host.getAttribute('Name') ?? '');
	if (hostsNode && (names.length === 0 || names.length !== elements(hostsNode).length))
		invalid('Hosts');
	unique(names, (name) => name, 'duplicate host');
	const hosts = hostsNode
		? names.flatMap((name) => {
				const product = XML_BASE_HOSTS.get(name);
				if (product === undefined) return invalid('host name');
				return product === null ? [] : [product];
			})
		: OFFICE_ADD_IN_HOSTS;
	if (!hosts.includes(options.host)) throw new AddInManifestError('host');
	const permission = PERMISSIONS.get(value(root, 'Permissions') ?? 'Restricted');
	if (!permission) return invalid('Permissions');
	const sourceLocation = one(one(root, 'DefaultSettings'), 'SourceLocation');
	const baseOrigins = new Set<string>();
	const pane = localize(sourceLocation, 'SourceLocation', (v) => {
		const entry = url(v, 'SourceLocation');
		baseOrigins.add(origin(entry));
		return entry;
	});
	const declaredOrigins = all(one(root, 'AppDomains'), 'AppDomain').map((node) =>
		declaredOrigin(node.textContent, 'AppDomain'),
	);
	const allowed = new Set([...baseOrigins, ...declaredOrigins]);
	const icons = ['IconUrl', 'HighResolutionIconUrl'].flatMap((field) => {
		const icon = one(root, field);
		return icon
			? [{ kind: 'remote' as const, url: localize(icon, field, (v) => url(v, field)) }]
			: [];
	});
	const support = one(root, 'SupportUrl');
	if (support) localize(support, 'SupportUrl', (v) => url(v, 'SupportUrl'));
	const requirements = xmlRequirements(one(root, 'Requirements'));
	const missing = missingRequirements(requirements, options);
	const issues: AddInIssue[] = missing.map((requirement) => ({ kind: 'requirement', requirement }));
	const overrides = elements(root).filter((child) => child.localName === 'VersionOverrides');
	if (overrides.length > 1) invalid('duplicate VersionOverrides');
	const override = overrides[0];
	const identityNode = override && one(override, 'WebApplicationInfo');
	const identity = identityNode
		? {
				id: guid(value(identityNode, 'Id'), 'WebApplicationInfo.Id'),
				resource: text(value(identityNode, 'Resource'), 'WebApplicationInfo.Resource'),
			}
		: null;
	if (identity) issues.push({ kind: 'unsupported', feature: 'web application identity' });
	const tabs = override ? readXmlOverrides(override, options, locale, allowed, issues) : [];
	return {
		format: 'xml',
		id,
		version: release,
		publisher,
		locale,
		hosts,
		permissions: [permission],
		name: localize(one(root, 'DisplayName'), 'DisplayName'),
		description: localize(one(root, 'Description'), 'Description'),
		requirements,
		entryOrigins: [...baseOrigins],
		declaredOrigins: [...new Set(declaredOrigins)],
		icons,
		identity,
		pane: missing.length ? null : pane,
		tabs: missing.length ? [] : tabs,
		issues,
	};
}
