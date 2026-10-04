// Styles live in real .css files next to their element and are imported as text (Vite's `?raw`).
declare module '*.css?raw' {
	const css: string;
	export default css;
}
