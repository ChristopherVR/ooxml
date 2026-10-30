import { z } from 'zod';

export const ManageCustomShowsSchema = z.object({
	filePath: z.string().describe('Path to the PPTX file'),
	action: z
		.enum([
			'list',
			'create',
			'rename',
			'setSlides',
			'addSlides',
			'removeSlides',
			'moveSlide',
			'copy',
			'delete',
		])
		.describe('Custom show action'),
	showId: z
		.string()
		.optional()
		.describe(
			'Custom show id from list (for rename, setSlides, addSlides, removeSlides, moveSlide, copy, delete)',
		),
	name: z.string().optional().describe('Show name (for create, rename, optionally copy)'),
	slideIndices: z
		.array(z.number().int().min(0))
		.optional()
		.describe('0-based slide indices in show order (for create, setSlides, addSlides)'),
	positions: z
		.array(z.number().int().min(0))
		.optional()
		.describe('0-based positions within the show list to remove (for removeSlides)'),
	from: z.number().int().min(0).optional().describe('Position to move from (for moveSlide)'),
	to: z.number().int().min(0).optional().describe('Position to move to (for moveSlide)'),
	index: z
		.number()
		.int()
		.min(0)
		.optional()
		.describe('Position within the show to insert at; default the end (for addSlides)'),
});
