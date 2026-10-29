// Prometheus text-format metrics and the structured logger of the reference server. No
// dependencies: counters are kept in memory, gauges are read when the endpoint is scraped.

const escapeLabel = (value) =>
	String(value).replace(/\\/gu, '\\\\').replace(/\n/gu, '\\n').replace(/"/gu, '\\"');

export function createMetrics() {
	/** @type {Map<string, { help: string, values: Map<string, number> }>} */
	const counters = new Map();
	/** @type {{ name: string, help: string, read: () => number | Array<{ labels: Record<string, string>, value: number }> }[]} */
	const gauges = [];
	const labelKey = (labels) =>
		Object.entries(labels)
			.map(([k, v]) => `${k}="${escapeLabel(v)}"`)
			.join(',');
	return {
		counter(name, help) {
			if (!counters.has(name)) counters.set(name, { help, values: new Map() });
			return (labels = {}, by = 1) => {
				const values = counters.get(name).values;
				const key = labelKey(labels);
				values.set(key, (values.get(key) ?? 0) + by);
			};
		},
		gauge(name, help, read) {
			gauges.push({ name, help, read });
		},
		/** The current value of a counter (for tests); sums every label set. */
		total(name) {
			return [...(counters.get(name)?.values.values() ?? [])].reduce((a, b) => a + b, 0);
		},
		render() {
			const lines = [];
			for (const [name, { help, values }] of counters) {
				lines.push(`# HELP ${name} ${help}`, `# TYPE ${name} counter`);
				if (values.size === 0) lines.push(`${name} 0`);
				for (const [key, value] of values) lines.push(`${name}${key ? `{${key}}` : ''} ${value}`);
			}
			for (const { name, help, read } of gauges) {
				lines.push(`# HELP ${name} ${help}`, `# TYPE ${name} gauge`);
				const value = read();
				if (typeof value === 'number') lines.push(`${name} ${value}`);
				else
					for (const item of value) lines.push(`${name}{${labelKey(item.labels)}} ${item.value}`);
			}
			return `${lines.join('\n')}\n`;
		},
	};
}

/**
 * A log function for the server. `text` writes the line as is; `json` writes one JSON object per
 * line with `time`, `level` (`warn` for lines starting with WARNING, else `info`) and `msg`.
 */
export function createLogger(format = 'text', write = (line) => console.log(line), now = Date.now) {
	if (format !== 'json') return (message) => write(message);
	return (message) => {
		const text = String(message);
		write(
			JSON.stringify({
				time: new Date(now()).toISOString(),
				level: /^WARNING\b/u.test(text) ? 'warn' : 'info',
				msg: text,
			}),
		);
	};
}
