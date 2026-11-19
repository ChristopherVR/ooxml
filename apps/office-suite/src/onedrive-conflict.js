const describe = (diff) => [
	...diff.changed.map((name) => `${name} (changed)`),
	...diff.added.map((name) => `${name} (added)`),
	...diff.removed.map((name) => `${name} (removed)`),
];

/**
 * The conflict dialog's Details: why combining stopped, then which package parts changed in
 * OneDrive and on this device. `comparison` is `{ remote, local }` from `diffOpcPackages`, or
 * `{ error }` when the packages could not be compared; the reasons are shown either way.
 */
export function conflictDetailsHtml(issues, comparison, escape) {
	const list = (title, names) =>
		names.length
			? `<p>${escape(title)}</p><ul>${names.map((name) => `<li><code>${escape(name)}</code></li>`).join('')}</ul>`
			: '';
	const why = list('Why the edits were not combined:', issues ?? []);
	if ('error' in comparison) {
		const { error } = comparison;
		const reason = error instanceof Error ? error.message : String(error);
		return `${why}<p>${escape(`Details are unavailable: ${reason}`)}</p>`;
	}
	const remote = describe(comparison.remote);
	const local = describe(comparison.local);
	return [
		why,
		remote.length
			? list('Changed in OneDrive since you opened the file:', remote)
			: '<p>OneDrive reported a newer version, but its content matches what you opened.</p>',
		list('Changed on this device:', local),
		list(
			'Changed in both:',
			remote.filter((name) => local.includes(name)),
		),
	].join('');
}
