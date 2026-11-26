import { elements, type XmlElement } from '../../xml/xml';
import type {
	AddInCommand,
	AddInIssue,
	AddInReadOptions,
	AddInRequirement,
	AddInTab,
} from './model';
import { identifier, invalid, missingRequirements, paneUrl, text, unique, url } from './validation';
import {
	all,
	one,
	OVERRIDE,
	TYPES,
	typeName,
	value,
	variants,
	XML_HOSTS,
	xmlRequirements,
} from './xml-fields';

export function readXmlOverrides(
	override: XmlElement,
	options: AddInReadOptions,
	defaultLocale: string,
	allowed: ReadonlySet<string>,
	issues: AddInIssue[],
	reviewRequirements: AddInRequirement[],
): AddInTab[] {
	if (override.namespaceURI !== OVERRIDE || typeName(override) !== 'VersionOverridesV1_0') {
		issues.push({ kind: 'unsupported', feature: 'VersionOverrides' });
		return [];
	}
	const requirements = xmlRequirements(one(override, 'Requirements'));
	const missing = missingRequirements(requirements, options);
	const resources = new Map<string, { kind: string; value: string }>();
	const ids = new Set<string>();
	for (const [container, tag] of [
		['Urls', 'Url'],
		['Images', 'Image'],
		['ShortStrings', 'String'],
		['LongStrings', 'String'],
	] as const) {
		for (const item of all(one(one(override, 'Resources'), container, TYPES), tag)) {
			const id = identifier(item.getAttribute('id'), 'resource ID');
			if (ids.has(id.toLowerCase())) invalid('duplicate XML resource ID');
			ids.add(id.toLowerCase());
			const convert =
				container === 'Urls'
					? (v: unknown) => paneUrl(v, id, allowed)
					: container === 'Images'
						? (v: unknown) => url(v, id)
						: undefined;
			resources.set(id, {
				kind: container,
				value: variants(item, convert)(options.locale, defaultLocale),
			});
		}
	}
	const resolve = (element: XmlElement | undefined, kind: string, field: string) => {
		const id = text(element?.getAttribute('resid'), field, 128);
		const resource = resources.get(id);
		if (!resource || resource.kind !== kind) return invalid(`unresolved ${field}`);
		return resource.value;
	};
	const readCommand = (
		control: XmlElement,
		issues: AddInIssue[],
		inherited?: AddInIssue,
	): AddInCommand => {
		const id = identifier(control.getAttribute('id'), 'control ID');
		const label = resolve(one(control, 'Label'), 'ShortStrings', 'control label');
		const tip = one(control, 'Supertip');
		if (tip) {
			resolve(one(tip, 'Title'), 'ShortStrings', 'supertip title');
			resolve(one(tip, 'Description'), 'LongStrings', 'supertip description');
		}
		for (const image of all(one(control, 'Icon'), 'Image', TYPES))
			resolve(image, 'Images', 'control image');
		const action = one(control, 'Action');
		if (!action) return invalid('control action');
		const type = typeName(action);
		if (type === 'ShowTaskpane') {
			const pane = resolve(one(action, 'SourceLocation'), 'Urls', 'task pane URL');
			const automatic: AddInIssue | undefined =
				value(action, 'TaskpaneId') === 'Office.AutoShowTaskpaneWithDocument'
					? { kind: 'unsupported', feature: 'automatic task pane' }
					: undefined;
			if (automatic) issues.push(automatic);
			const issue = inherited ?? automatic;
			return issue
				? { id, label, kind: 'inactive', issue }
				: { id, label, kind: 'taskPane', url: pane };
		}
		const issue: AddInIssue = { kind: 'unsupported', feature: type };
		issues.push(issue);
		return { id, label, kind: 'inactive', issue };
	};
	const readTab = (tab: XmlElement, issues: AddInIssue[], inherited?: AddInIssue): AddInTab => {
		const id = identifier(tab.getAttribute('id'), 'tab ID');
		const builtIn = tab.localName === 'OfficeTab';
		const builtInIssue: AddInIssue | undefined = builtIn
			? { kind: 'unsupported', feature: 'built-in ribbon tab' }
			: undefined;
		if (builtInIssue) issues.push(builtInIssue);
		const issue = inherited ?? builtInIssue;
		const groups = all(tab, 'Group').map((group) => {
			const commands = all(group, 'Control').flatMap((control) => {
				const type = typeName(control);
				if (type === 'Button') return [readCommand(control, issues, issue)];
				if (type === 'Menu') {
					resolve(one(control, 'Label'), 'ShortStrings', 'menu label');
					return all(one(control, 'Items'), 'Item').map((item) => readCommand(item, issues, issue));
				}
				issues.push({ kind: 'unsupported', feature: type });
				return [];
			});
			unique(commands, (command) => command.id, 'duplicate command ID');
			return {
				id: identifier(group.getAttribute('id'), 'group ID'),
				label: resolve(one(group, 'Label'), 'ShortStrings', 'group label'),
				commands,
			};
		});
		unique(groups, (group) => group.id, 'duplicate group ID');
		return {
			id,
			label: builtIn ? id : resolve(one(tab, 'Label'), 'ShortStrings', 'tab label'),
			groups,
		};
	};
	const hosts = all(one(override, 'Hosts'), 'Host');
	unique(hosts, typeName, 'duplicate override host');
	const editorHosts = hosts.flatMap((host) => {
		const product = XML_HOSTS.get(typeName(host));
		if (product === undefined) return invalid('override host type');
		return product === null ? [] : [{ host, product }];
	});
	const requirement = missing[0];
	const inactive = requirement ? ({ kind: 'requirement', requirement } as const) : undefined;
	const tabs: AddInTab[] = [];
	for (const { host, product } of editorHosts) {
		const applies = product === options.host;
		const hostIssues: AddInIssue[] = [];
		const desktop = one(host, 'DesktopFormFactor');
		const functionFile = one(desktop, 'FunctionFile');
		if (functionFile) resolve(functionFile, 'Urls', 'function file');
		for (const point of all(desktop, 'ExtensionPoint')) {
			const type = typeName(point);
			if (type !== 'PrimaryCommandSurface') {
				hostIssues.push({ kind: 'unsupported', feature: type });
				continue;
			}
			const result = elements(point)
				.filter(
					(tab) =>
						tab.namespaceURI === OVERRIDE && ['CustomTab', 'OfficeTab'].includes(tab.localName),
				)
				.map((tab) => readTab(tab, hostIssues, inactive));
			if (!result.length || result.some((tab) => !tab.groups.length))
				invalid('ribbon tabs and groups');
			if (applies) tabs.push(...result);
		}
		if (applies) {
			reviewRequirements.push(...requirements);
			for (const requirement of missing) issues.push({ kind: 'requirement', requirement });
			issues.push(...hostIssues);
		}
	}
	unique(tabs, (tab) => tab.id, 'duplicate tab ID');
	return tabs;
}
