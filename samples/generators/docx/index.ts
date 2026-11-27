/** The Word samples. */
import type { SampleDefinition } from '../lib/types';
import { servicesAgreementForm } from './agreement/index';
import { communityNewsletter } from './newsletter/index';
import { technicalReport } from './report/index';

export const samples: SampleDefinition[] = [
	technicalReport,
	communityNewsletter,
	servicesAgreementForm,
];
