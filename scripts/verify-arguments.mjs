/** Keep shell commands below Windows command-line limits, preserving each literal argument. */
export function argumentBatches(prefix, args, limit = 4000) {
	const commands = [];
	let command = prefix;
	for (const arg of args) {
		const quoted = `'${arg.replaceAll("'", "'\\''")}'`;
		if (prefix.length + quoted.length + 1 > limit)
			throw new Error(`Verification argument exceeds command limit: ${arg}`);
		if (command.length + quoted.length + 1 > limit) {
			commands.push(command);
			command = prefix;
		}
		command += ` ${quoted}`;
	}
	if (command !== prefix) commands.push(command);
	return commands;
}
