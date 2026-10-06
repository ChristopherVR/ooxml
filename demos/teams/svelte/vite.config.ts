import { svelte } from '@sveltejs/vite-plugin-svelte';
import { defineConfig } from 'vite';
import { sharedConfig } from '../vite.shared.ts';

export default defineConfig({ ...sharedConfig(5177), plugins: [svelte()] });
