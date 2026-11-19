/**
 * Chromium's accessibility tree, the one a screen reader reads, flattened to the nodes it exposes.
 */
import type { Page } from '@playwright/test';

export interface AxNode {
	role: string;
	name: string;
	properties: Record<string, unknown>;
}

interface RawNode {
	nodeId: string;
	ignored?: boolean;
	role?: { value?: string };
	name?: { value?: string };
	properties?: { name: string; value?: { value?: unknown } }[];
}

/** Every node the tree exposes, in document order. */
export async function axTree(page: Page): Promise<AxNode[]> {
	const cdp = await page.context().newCDPSession(page);
	try {
		await cdp.send('Accessibility.enable');
		const { nodes } = (await cdp.send('Accessibility.getFullAXTree')) as { nodes: RawNode[] };
		return nodes
			.filter((node) => !node.ignored && node.role?.value && node.role.value !== 'none')
			.map((node) => ({
				role: String(node.role?.value),
				name: String(node.name?.value ?? '').trim(),
				properties: Object.fromEntries(
					(node.properties ?? []).map((property) => [property.name, property.value?.value]),
				),
			}));
	} finally {
		await cdp.detach();
	}
}

/** Nodes with `role` whose name matches `name`. */
export const nodes = (tree: readonly AxNode[], role: string, name?: string | RegExp) =>
	tree.filter(
		(node) =>
			node.role === role &&
			(name === undefined ||
				(typeof name === 'string' ? node.name === name : name.test(node.name))),
	);

/** A short listing of the tree, for failure messages. */
export const outline = (tree: readonly AxNode[], roles: readonly string[]) =>
	tree
		.filter((node) => roles.includes(node.role))
		.map((node) => `${node.role} "${node.name.slice(0, 50)}"`)
		.slice(0, 40)
		.join('\n');
