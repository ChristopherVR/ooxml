/** The Visio sample drawings. */
import type { SampleDefinition } from '../lib/types';
import { architecture } from './architecture/index';
import { flowchart } from './flowchart/index';
import { orgChart } from './orgchart/index';

export const samples: SampleDefinition[] = [flowchart, architecture, orgChart];
