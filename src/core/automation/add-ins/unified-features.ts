import type {
	AddInCommand,
	AddInIssue,
	AddInReadOptions,
	AddInRequirement,
	AddInTab,
	OfficeAddInHost,
} from './model';
import { OFFICE_ADD_IN_HOSTS } from './model';
import { object, optionalArray, optionalObject, records } from './json';
import {
	identifier,
	invalid,
	missingRequirements,
	origin,
	paneUrl,
	requirement,
	text,
	unique,
	url,
} from './validation';

const SCOPES = new Map<string, OfficeAddInHost>([
	['document', 'word'],
	['workbook', 'excel'],
	['presentation', 'powerpoint'],
]);

export function featureRequirements(input: unknown, options: AddInReadOptions) {
	const entry = optionalObject(input, 'requirements');
	const scopes = optionalArray(entry.scopes, 'requirements.scopes').map((scope) => {
		const name = text(scope, 'scope');
		if (!SCOPES.has(name) && name !== 'mail') invalid('host scope');
		return name;
	});
	unique(scopes, (scope) => scope, 'duplicate scope');
	if (entry.scopes !== undefined && !scopes.length) invalid('empty scopes');
	const factors = optionalArray(entry.formFactors, 'requirements.formFactors');
	if (factors.some((factor) => factor !== 'mobile' && factor !== 'desktop')) invalid('form factor');
	const requirements = records(entry.capabilities, 'requirements.capabilities').map((capability) =>
		requirement(capability.name, capability.minVersion, capability.maxVersion),
	);
	unique(requirements, (item) => item.name, 'duplicate capability');
	const hosts = scopes.length
		? OFFICE_ADD_IN_HOSTS.filter((host) => scopes.some((scope) => SCOPES.get(scope) === host))
		: OFFICE_ADD_IN_HOSTS;
	const missing = missingRequirements(requirements, options)[0];
	let issue: AddInIssue | undefined;
	if (!hosts.includes(options.host)) issue = { kind: 'unsupported', feature: scopes.join(', ') };
	else if (factors.length && !factors.includes('desktop'))
		issue = { kind: 'unsupported', feature: 'mobile form factor' };
	else if (missing) issue = { kind: 'requirement', requirement: missing };
	return { hosts, requirements, issue };
}

type Action = { kind: 'taskPane'; url: string } | { kind: 'inactive'; issue: AddInIssue };

type JsonRecord = ReturnType<typeof object>;
type ExtensionContext = {
	options: AddInReadOptions;
	declaredOrigins: readonly string[];
	entryOrigins: Set<string>;
	permissionIssue?: AddInIssue;
	icon: (value: unknown) => void;
	label: (path: string, value: unknown, max?: number) => string;
	issues: AddInIssue[];
};

function readActions(
	extension: JsonRecord,
	context: ExtensionContext,
	inherited: AddInIssue | undefined,
) {
	const { options, declaredOrigins, entryOrigins, issues } = context;
	const actions = new Map<string, Action>();
	const runtimes = records(extension.runtimes, 'runtimes');
	unique(runtimes, (runtime) => identifier(runtime.id, 'runtime ID'), 'duplicate runtime ID');
	const ids = new Set<string>();
	for (const runtime of runtimes) {
		const { issue: requirements } = featureRequirements(runtime.requirements, options);
		const type = identifier(runtime.type, 'runtime type');
		const lifetime = runtime.lifetime ?? 'short';
		if (lifetime !== 'long' && lifetime !== 'short') invalid('runtime lifetime');
		const code = object(runtime.code, 'runtime code');
		const page = code.page === undefined ? null : url(code.page, 'runtime page');
		if (page) entryOrigins.add(origin(page));
		if (code.script !== undefined) {
			const allowed = new Set(page ? [origin(page), ...declaredOrigins] : declaredOrigins);
			if (page || allowed.size) paneUrl(code.script, 'runtime script', allowed);
			else url(code.script, 'runtime script');
		}
		const unsupported =
			type !== 'general' ? type : lifetime === 'long' ? 'shared runtime' : undefined;
		for (const action of records(runtime.actions, 'runtime actions')) {
			const id = identifier(action.id, 'action ID');
			if (ids.has(id.toLowerCase())) invalid('duplicate action ID');
			ids.add(id.toLowerCase());
			const actionType = identifier(action.type, 'action type');
			const issue =
				inherited ??
				requirements ??
				(unsupported || actionType !== 'openPage'
					? ({ kind: 'unsupported', feature: unsupported ?? actionType } as const)
					: undefined);
			if (actionType === 'openPage' && !page) invalid('openPage runtime page');
			if (issue) {
				issues.push(issue);
				actions.set(id, { kind: 'inactive', issue });
			} else if (page) {
				actions.set(id, { kind: 'taskPane', url: page });
			}
		}
	}
	return actions;
}

function readRibbon(
	ribbon: JsonRecord,
	path: string,
	actions: ReadonlyMap<string, Action>,
	context: ExtensionContext,
	inherited: AddInIssue | undefined,
): AddInTab[] {
	const { icon, label, options, issues } = context;
	const contexts = optionalArray(ribbon.contexts, 'ribbon contexts').map((entry) =>
		identifier(entry, 'ribbon context'),
	);
	const requirements = featureRequirements(ribbon.requirements, options);
	const inactive =
		inherited ??
		requirements.issue ??
		(contexts.length && !contexts.includes('default')
			? ({ kind: 'unsupported', feature: contexts.join(', ') } as const)
			: undefined);
	if (inactive) issues.push(inactive);
	const icons = (entry: JsonRecord) => {
		for (const ref of records(entry.icons, 'ribbon icons')) icon(ref.url);
	};
	const readControl = (control: JsonRecord, field: string, issue?: AddInIssue): AddInCommand[] => {
		const id = identifier(control.id, 'control ID');
		const type = identifier(control.type, 'control type');
		const name = label(`${field}.label`, control.label, 64);
		icons(control);
		if (type === 'menu') {
			return records(control.items, 'menu items').flatMap((item, index) => {
				if (item.type === 'menu') invalid('nested menu');
				return readControl(item, `${field}.items[${index}]`, issue);
			});
		}
		if (type !== 'button' && type !== 'menuItem')
			return [{ id, label: name, kind: 'inactive', issue: { kind: 'unsupported', feature: type } }];
		const action = actions.get(identifier(control.actionId, 'action reference'));
		if (!action) return invalid('unresolved ribbon action');
		const unavailable =
			issue ??
			(control.enabled === false || control.overriddenByRibbonApi === true
				? ({ kind: 'unsupported', feature: 'dynamic ribbon command' } as const)
				: undefined);
		return [
			{
				id,
				label: name,
				...(unavailable ? ({ kind: 'inactive', issue: unavailable } as const) : action),
			},
		];
	};
	return records(ribbon.tabs, 'ribbon tabs').map((tab, index): AddInTab => {
		const tabPath = `${path}.tabs[${index}]`;
		const builtIn = tab.builtInTabId !== undefined;
		if (builtIn && (tab.id !== undefined || tab.customTabId !== undefined))
			invalid('ambiguous tab ID');
		const id = identifier(builtIn ? tab.builtInTabId : (tab.id ?? tab.customTabId), 'tab ID');
		const issue = builtIn
			? ({ kind: 'unsupported', feature: 'built-in ribbon tab' } as const)
			: inactive;
		if (builtIn) issues.push({ kind: 'unsupported', feature: 'built-in ribbon tab' });
		const groups = records(tab.groups, 'ribbon groups').map((group, groupIndex) => {
			const groupPath = `${tabPath}.groups[${groupIndex}]`;
			icons(group);
			const commands = records(group.controls, 'ribbon controls').flatMap((control, controlIndex) =>
				readControl(control, `${groupPath}.controls[${controlIndex}]`, issue),
			);
			unique(commands, (command) => command.id, 'duplicate command ID');
			return {
				id: identifier(group.id, 'group ID'),
				label: label(`${groupPath}.label`, group.label, 64),
				commands,
			};
		});
		if (!groups.length) invalid('ribbon groups');
		unique(groups, (group) => group.id, 'duplicate group ID');
		return { id, label: builtIn ? id : label(`${tabPath}.label`, tab.label, 64), groups };
	});
}

export function readExtensions(input: unknown, context: ExtensionContext) {
	const { options, issues } = context;
	const hosts = new Set<OfficeAddInHost>();
	const requirements: AddInRequirement[] = [];
	const tabs: AddInTab[] = [];
	const extensionKeys = new Set([
		'requirements',
		'runtimes',
		'ribbons',
		'getStartedMessages',
		'alternates',
	]);
	for (const [index, extension] of records(input, 'Office extensions').entries()) {
		const condition = featureRequirements(extension.requirements, options);
		for (const host of condition.hosts) hosts.add(host);
		const applies = condition.hosts.includes(options.host);
		if (applies) requirements.push(...condition.requirements);
		if (condition.issue) issues.push(condition.issue);
		const actions = readActions(extension, context, condition.issue ?? context.permissionIssue);
		for (const [key, value] of Object.entries(extension)) {
			if (!extensionKeys.has(key) && value !== undefined)
				issues.push({ kind: 'unsupported', feature: identifier(key, 'extension feature') });
		}
		for (const [ribbonIndex, ribbon] of records(extension.ribbons, 'ribbons').entries()) {
			const entries = readRibbon(
				ribbon,
				`extensions[${index}].ribbons[${ribbonIndex}]`,
				actions,
				context,
				condition.issue,
			);
			if (applies) tabs.push(...entries);
		}
	}
	return { hosts: [...hosts], requirements, tabs };
}
