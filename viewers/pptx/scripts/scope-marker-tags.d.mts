export declare const SCOPE_MARKERS: string[];
export declare function isCustomElementTag(name: string): boolean;
export declare function findScopedCustomElements(
	source: string,
): { tag: string; marker: string; line: number }[];
export declare function sourceFiles(dir: string, extensions: readonly string[]): string[];
export declare function scopedCustomElementsIn(
	dir: string,
	extensions: readonly string[],
): string[];
