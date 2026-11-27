/** The Excel sample workbooks. */
import type { SampleDefinition } from '../lib/types';
import { financialModel } from './financial-model/index';
import { operationsTracker } from './operations-tracker/index';
import { salesAnalysis } from './sales-analysis/index';

export const samples: SampleDefinition[] = [financialModel, salesAnalysis, operationsTracker];
