import { describe, expect, it } from 'vitest';
import {
	fileTextChunks,
	readRecipientFile,
	RECIPIENT_LIMITS,
	snapshotRecipients,
} from './recipient-data';

async function collect(chunks: AsyncIterable<string>): Promise<string[]> {
	const out: string[] = [];
	for await (const chunk of chunks) out.push(chunk);
	return out;
}

describe('local recipient files', () => {
	it('decodes UTF-8 slices without splitting a character', async () => {
		const text = 'Name\nZoë\n東京\n';
		const chunks = await collect(fileTextChunks(new Blob([text]), 3));
		expect(chunks.length).toBeGreaterThan(3);
		expect(chunks.join('')).toBe(text);
	});

	it('reads a CSV larger than the old 1,000 record cap in slices', async () => {
		const rows = Array.from(
			{ length: 5_000 },
			(_, i) => `Person ${i},${String(i).padStart(5, '0')}`,
		);
		const file = new File([`Name,Postal\n${rows.join('\n')}\n`], 'people.csv');
		const data = await readRecipientFile(file);
		expect(data.fields).toEqual(['Name', 'Postal']);
		expect(data.records).toHaveLength(5_000);
		expect(data.records[4_999]).toEqual({ Name: 'Person 4999', Postal: '04999' });
	});

	it('refuses files and lists over the documented caps', async () => {
		const big = { size: RECIPIENT_LIMITS.delimitedBytes + 1, name: 'big.csv' } as File;
		await expect(readRecipientFile(big)).rejects.toThrow(RangeError);
		const workbook = { size: RECIPIENT_LIMITS.workbookBytes + 1, name: 'big.xlsx' } as File;
		await expect(readRecipientFile(workbook)).rejects.toThrow(RangeError);
		const wide = Array.from({ length: RECIPIENT_LIMITS.fields + 1 }, (_, i) => `F${i}`);
		await expect(readRecipientFile(new File([`${wide.join(',')}\n`], 'wide.csv'))).rejects.toThrow(
			RangeError,
		);
		await expect(readRecipientFile(new File(['Bad\u0001Name\nx\n'], 'bad.csv'))).rejects.toThrow(
			TypeError,
		);
		const record = Object.fromEntries(Array.from({ length: 50 }, (_, i) => [`F${i}`, 'x']));
		const many = Array.from({ length: RECIPIENT_LIMITS.cells / 50 + 1 }, () => record);
		expect(() => snapshotRecipients({ records: many })).toThrow(RangeError);
	});
});
