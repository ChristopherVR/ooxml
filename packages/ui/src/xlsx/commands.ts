// Command registry: every ribbon button, menu item and shortcut runs a command by id. The
// implementations live in `commands/` (UI-COMMANDS) and the grid; this file is only the registry.
import type { EditorContext } from './context';

export interface Command {
	/** Stable kebab id: 'home.bold', 'home.merge-center', 'view.freeze-panes'. */
	id: string;
	/** English text (translation key). */
	label: string;
	/** An ooxml-ui icon name or a local icon registered with `registerIcon`. */
	icon?: string;
	/** Display text, for example 'Ctrl+B'. */
	shortcut?: string;
	/** True: disabled in read-only mode. */
	editing?: boolean;
	enabled?(ctx: EditorContext): boolean;
	/** Toggle state (bold of the active cell, ...). */
	checked?(ctx: EditorContext): boolean;
	/** Value for select-like controls (font name, size, number format). */
	value?(ctx: EditorContext): string;
	run(ctx: EditorContext, arg?: unknown): void | Promise<void>;
}

export interface CommandRegistry {
	register(command: Command): void;
	registerAll(commands: Command[]): void;
	get(id: string): Command | undefined;
	list(): Command[];
	/** False when the command is missing or disabled; emits 'ribbon-action'. */
	run(id: string, arg?: unknown): Promise<boolean>;
	isEnabled(id: string): boolean;
}

export interface CommandRegistryOptions {
	/** Called when a command throws, so the shell can show a toast. */
	onError?(id: string, error: unknown): void;
	/** Ids hidden by the host (`hiddenActions`) never run. */
	isHidden?(id: string): boolean;
}

export function createCommandRegistry(
	context: () => EditorContext,
	options: CommandRegistryOptions = {},
): CommandRegistry {
	const commands = new Map<string, Command>();
	const isEnabled = (id: string): boolean => {
		const command = commands.get(id);
		if (!command || options.isHidden?.(id)) return false;
		const ctx = context();
		if (command.editing && ctx.readOnly()) return false;
		try {
			return command.enabled ? command.enabled(ctx) : true;
		} catch {
			return false;
		}
	};
	return {
		register(command) {
			commands.set(command.id, command);
		},
		registerAll(list) {
			for (const command of list) commands.set(command.id, command);
		},
		get: (id) => commands.get(id),
		list: () => [...commands.values()],
		isEnabled,
		async run(id, arg) {
			const command = commands.get(id);
			if (!command || !isEnabled(id)) return false;
			const ctx = context();
			ctx.emit('ribbon-action', { id });
			try {
				await command.run(ctx, arg);
			} catch (error) {
				options.onError?.(id, error);
				return false;
			}
			ctx.requestRender();
			return true;
		},
	};
}
