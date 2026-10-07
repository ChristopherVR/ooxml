import { createToggleFormat } from 'ooxml-core/docx/ui';
import { styleModelOf } from './run-styles';
import { schema } from './schema';

export type { ToggleKey } from 'ooxml-core/docx/ui';
export const toggleFormat = createToggleFormat(schema, styleModelOf);
