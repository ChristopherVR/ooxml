import { defineConfig } from 'vite';
import { sharedConfig } from '../vite.shared.ts';

export default defineConfig({ ...sharedConfig(5178), oxc: { decorator: { legacy: true } } });
