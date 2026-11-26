import { safeZipPath } from '../../opc/zip-headers';
import type { AddInIcon, AddInIssue, AddInManifest, AddInReadOptions } from './model';
import { object, optionalObject, parseManifestJson, records, optionalArray } from './json';
import { readExtensions } from './unified-features';
import {
	AddInManifestError,
	declaredOrigin,
	guid,
	identifier,
	invalid,
	localeCandidates,
	localeTag,
	PERMISSIONS,
	text,
	unique,
	url,
	version,
} from './validation';

export interface ManifestAssets {
	has(path: string): boolean;
	readText(path: string): Promise<string>;
}

// Microsoft 365 app manifest releases with the documented Office task-pane extension fields.
const SCHEMA_VERSIONS = new Set(['1.25', '1.26', '1.27', '1.28', '1.29', '1.30']);
const INACTIVE_APP_FEATURES = [
	'bots',
	'composeExtensions',
	'configurableTabs',
	'staticTabs',
	'connectors',
	'webApplicationInfo',
	'copilotAgents',
	'meetingExtensionDefinition',
	'activityTypes',
	'dashboardCards',
];

async function loadTranslations(raw: unknown, assets: ManifestAssets, options: AddInReadOptions) {
	const info = optionalObject(raw, 'localizationInfo');
	const locale =
		info.defaultLanguageTag === undefined ? 'en-us' : localeTag(info.defaultLanguageTag);
	const translations = new Map<string, ReturnType<typeof object>>();
	for (const entry of records(info.additionalLanguages, 'additionalLanguages')) {
		const tag = localeTag(entry.languageTag);
		if (translations.has(tag) || tag === locale) invalid('duplicate locale');
		const path = packagePath(entry.file);
		const data = object(parseManifestJson(await assets.readText(path)), 'locale file');
		translations.set(tag, {
			...optionalObject(data.localizationKeys, 'localizationKeys'),
			...data,
		});
	}
	const candidates = [
		...new Set([...localeCandidates(options.locale), ...localeCandidates(locale)]),
	];
	const label = (key: string, fallback: unknown, limit = 4096): string => {
		const original = text(fallback, key, limit);
		for (const candidate of candidates) {
			const value = translations.get(candidate)?.[key];
			if (value !== undefined) return text(value, key, limit);
		}
		return original;
	};
	return { locale, label };
}

function packagePath(value: unknown): string {
	return safeZipPath(text(value, 'package resource', 240), false, () => {
		throw new AddInManifestError('package', 'resource path');
	});
}

export async function readUnifiedManifest(
	source: string,
	assets: ManifestAssets,
	options: AddInReadOptions,
): Promise<AddInManifest> {
	const root = object(parseManifestJson(source), 'manifest');
	if (typeof root.manifestVersion !== 'string' || !SCHEMA_VERSIONS.has(root.manifestVersion))
		throw new AddInManifestError('version', 'manifestVersion');
	const id = guid(root.id, 'id');
	const release = version(root.version, 'version', /^\d{1,5}(?:\.\d{1,5}){2}$/);
	const developer = object(root.developer, 'developer');
	const publisher = text(developer.name, 'developer.name', 256);
	for (const field of ['websiteUrl', 'privacyUrl', 'termsOfUseUrl']) {
		if (developer[field] !== undefined) url(developer[field], field);
	}
	const declaredOrigins = [
		...new Set(
			optionalArray(root.validDomains, 'validDomains').map((domain) =>
				declaredOrigin(
					domain,
					'validDomains',
					typeof domain === 'string' && !domain.includes('://'),
				),
			),
		),
	];
	const { locale, label } = await loadTranslations(root.localizationInfo, assets, options);
	const issues: AddInIssue[] = [];
	const authorization = optionalObject(root.authorization, 'authorization');
	let permissionIssue: AddInIssue | undefined;
	const permissions = records(
		optionalObject(authorization.permissions, 'permissions').resourceSpecific,
		'resourceSpecific',
	).map((permission) => {
		const name = identifier(permission.name, 'permission name');
		if (permission.type !== 'Delegated' && permission.type !== 'Application')
			invalid('permission type');
		if (permission.type !== 'Delegated' || ![...PERMISSIONS.values()].includes(name)) {
			const issue: AddInIssue = { kind: 'unsupported', feature: name };
			permissionIssue ??= issue;
			issues.push(issue);
		}
		return name;
	});
	unique(permissions, (name) => name, 'duplicate permission');
	if (permissions.filter((permission) => permission.startsWith('Document.')).length > 1)
		invalid('ambiguous document permission');
	const icons = new Map<string, AddInIcon>();
	const icon = (value: unknown, packageOnly = false) => {
		const source = text(value, 'icon source');
		if (!packageOnly && /^[a-z][a-z0-9+.-]*:/i.test(source)) {
			const target = url(source, 'icon URL');
			icons.set(target, { kind: 'remote', url: target });
		} else {
			const path = packagePath(source);
			if (!assets.has(path)) throw new AddInManifestError('package', 'missing icon');
			icons.set(path, { kind: 'package', path });
		}
	};
	if (root.icons !== undefined) {
		const appIcons = object(root.icons, 'icons');
		icon(appIcons.color, true);
		icon(appIcons.outline, true);
	}
	const identity =
		root.webApplicationInfo === undefined
			? null
			: (() => {
					const info = object(root.webApplicationInfo, 'webApplicationInfo');
					const id = guid(info.id, 'webApplicationInfo.id');
					return info.resource === undefined
						? { id }
						: { id, resource: text(info.resource, 'webApplicationInfo.resource') };
				})();
	for (const feature of INACTIVE_APP_FEATURES) {
		const value = root[feature];
		if (value !== undefined && (!Array.isArray(value) || value.length))
			issues.push({ kind: 'unsupported', feature });
	}
	const entryOrigins = new Set<string>();
	const { hosts, requirements, tabs } = readExtensions(root.extensions, {
		options,
		declaredOrigins,
		entryOrigins,
		icon,
		label,
		issues,
		...(permissionIssue && { permissionIssue }),
	});
	if (!hosts.length) throw new AddInManifestError('extension', 'Office task pane');
	if (!hosts.includes(options.host)) throw new AddInManifestError('host');
	const pane =
		tabs
			.flatMap((tab) => tab.groups.flatMap((group) => group.commands))
			.find((command) => command.kind === 'taskPane')?.url ?? null;
	return {
		format: 'unified',
		id,
		version: release,
		publisher,
		locale,
		hosts,
		requirements,
		permissions,
		declaredOrigins,
		entryOrigins: [...entryOrigins],
		identity,
		pane,
		tabs,
		issues,
		icons: [...icons.values()],
		name: label('name.short', object(root.name, 'name').short, 30),
		description: label('description.short', object(root.description, 'description').short, 80),
	};
}
