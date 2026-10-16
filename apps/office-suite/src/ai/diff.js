/**
 * Before-and-after views of proposed edits. `describeChanges` reads the old text and values from
 * the context the request sent, so a card can show exactly what each edit replaces; `wordDiff`
 * marks the words removed and added.
 */

/** Old and new text for each proposed change, resolved against the files the request read. */
export function describeChanges(changes, files) {
	return changes.map((change) => {
		const content = files.find((f) => f.id === change.documentId)?.content;
		if (change.type === 'word_run') {
			const paragraph = content?.paragraphs?.find((p) => p.id === change.paragraphId);
			const runs = paragraph?.runs.map((r) => r.text) ?? [];
			const after = runs.map((text, i) => (i === change.runIndex ? change.text : text));
			return { where: 'Paragraph', before: runs.join(''), after: after.join('') };
		}
		if (change.type === 'slide_run') {
			const element = content?.slides?.[change.slideIndex]?.elements.find(
				(e) => e.id === change.elementId,
			);
			const runs = element?.runs?.map((r) => r.text) ?? [];
			const after = runs.map((text, i) => (i === change.runIndex ? change.text : text));
			return {
				where: `Slide ${change.slideIndex + 1}`,
				before: runs.join(''),
				after: after.join(''),
			};
		}
		if (change.type === 'shape_text') {
			const page = content?.pages?.find((p) => p.id === change.pageId);
			const find = (shapes) => {
				for (const shape of shapes ?? []) {
					if (shape.id === change.shapeId) return shape;
					const child = find(shape.children);
					if (child) return child;
				}
			};
			return {
				where: page ? `Page ${page.name}` : 'Shape',
				before: find(page?.shapes)?.text ?? '',
				after: change.text,
			};
		}
		const sheet = content?.sheets?.find((s) => s.index === change.sheetIndex);
		return {
			where: sheet?.name ?? `Sheet ${change.sheetIndex + 1}`,
			cells: (change.cells ?? []).map((cell) => {
				const old = sheet?.cells.find((c) => c.address === cell.address.toUpperCase());
				return {
					address: cell.address,
					before: old ? (old.formula ? `=${old.formula}` : String(old.value ?? '')) : '',
					after: cell.input,
				};
			}),
		};
	});
}

const tokens = (text) => String(text).match(/\s+|[^\s]+/gu) ?? [];

/** A fragment showing `before` turned into `after`, with removed and added words marked. */
export function wordDiff(before, after) {
	const a = tokens(before),
		b = tokens(after);
	const out = document.createDocumentFragment();
	const push = (tag, text) => {
		if (!text) return;
		const last = out.lastChild;
		if (last && (last.nodeName.toLowerCase() === tag || (tag === '#text' && last.nodeType === 3)))
			last.textContent += text;
		else
			out.append(
				tag === '#text'
					? document.createTextNode(text)
					: Object.assign(document.createElement(tag), { textContent: text }),
			);
	};
	// Long texts skip the table and show a plain replacement.
	if (a.length * b.length > 250_000) {
		push('del', before);
		push('ins', after);
		return out;
	}
	const table = Array.from({ length: a.length + 1 }, () => new Uint16Array(b.length + 1));
	for (let i = a.length - 1; i >= 0; i--)
		for (let j = b.length - 1; j >= 0; j--)
			table[i][j] =
				a[i] === b[j] ? table[i + 1][j + 1] + 1 : Math.max(table[i + 1][j], table[i][j + 1]);
	let i = 0,
		j = 0;
	while (i < a.length || j < b.length) {
		if (i < a.length && j < b.length && a[i] === b[j]) {
			push('#text', a[i]);
			i++;
			j++;
		} else if (i < a.length && (j === b.length || table[i + 1][j] >= table[i][j + 1]))
			push('del', a[i++]);
		else push('ins', b[j++]);
	}
	return out;
}
