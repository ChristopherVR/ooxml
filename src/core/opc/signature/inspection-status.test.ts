import { describe, it, expect } from 'vitest';

import { computeDetailStatus, computeVerificationStatus } from './inspection-status';
import type {
	CertificateRevocationStatus,
	DigitalSignatureVerificationStatus,
	SignatureDetail,
	SignatureDetailStatus,
	SignatureReferenceCheck,
	SignatureValidationPolicy,
	TimestampAuthorityStatus,
} from './types';

const defaultPolicy: SignatureValidationPolicy = {
	requireRevocationCheck: false,
	failOnRevocationUnknown: false,
	requireTimestamp: false,
};

/** Create a mock reference check with defaults. */
function mockReferenceCheck(
	overrides: Partial<SignatureReferenceCheck> = {},
): SignatureReferenceCheck {
	return {
		uri: '/ppt/slides/slide1.xml',
		existsInPackage: true,
		digestStatus: 'verified',
		transformAlgorithms: [],
		...overrides,
	};
}

/**
 * Create a mock detail object suitable for `computeDetailStatus`.
 * All fields default to a "healthy" state that would produce 'verified'.
 */
function mockDetail(
	overrides: Partial<
		Pick<
			SignatureDetail,
			| 'signatureValueStatus'
			| 'missingPartReferences'
			| 'referenceChecks'
			| 'coverageGaps'
			| 'certificateTrustStatus'
			| 'certificateRevocationStatus'
			| 'timestampAuthorityStatus'
		>
	> = {},
): Pick<
	SignatureDetail,
	| 'signatureValueStatus'
	| 'missingPartReferences'
	| 'referenceChecks'
	| 'coverageGaps'
	| 'certificateTrustStatus'
	| 'certificateRevocationStatus'
	| 'timestampAuthorityStatus'
> {
	return {
		signatureValueStatus: 'verified',
		missingPartReferences: [],
		referenceChecks: [mockReferenceCheck()],
		coverageGaps: [],
		certificateTrustStatus: 'trusted',
		certificateRevocationStatus: 'good',
		timestampAuthorityStatus: 'valid',
		...overrides,
	};
}

/** Create a full SignatureDetail suitable for `computeVerificationStatus`. */
function mockFullDetail(overrides: Partial<SignatureDetail> = {}): SignatureDetail {
	return {
		path: '_xmlsignatures/sig1.xml',
		referenceCount: 1,
		missingPartReferences: [],
		unsupportedTransforms: [],
		referenceChecks: [mockReferenceCheck()],
		uncoveredParts: [],
		coverageGaps: [],
		signatureValueStatus: 'verified',
		certificateTrustStatus: 'trusted',
		certificateRevocationStatus: 'good',
		timestampAuthorityStatus: 'valid',
		status: 'verified',
		...overrides,
	};
}

describe('computeDetailStatus', () => {
	it('returns signature-invalid when signatureValueStatus is invalid', () => {
		const detail = mockDetail({ signatureValueStatus: 'invalid' });
		expect(computeDetailStatus(detail, defaultPolicy)).toBe('signature-invalid');
	});

	it('returns reference-missing when missingPartReferences is non-empty', () => {
		const detail = mockDetail({
			missingPartReferences: ['ppt/slides/slide2.xml'],
		});
		expect(computeDetailStatus(detail, defaultPolicy)).toBe('reference-missing');
	});

	it('returns digest-mismatch when a reference check has digestStatus mismatch', () => {
		const detail = mockDetail({
			referenceChecks: [mockReferenceCheck({ digestStatus: 'mismatch' })],
		});
		expect(computeDetailStatus(detail, defaultPolicy)).toBe('digest-mismatch');
	});

	it('returns certificate-revoked when certificateRevocationStatus is revoked', () => {
		const detail = mockDetail({ certificateRevocationStatus: 'revoked' });
		expect(computeDetailStatus(detail, defaultPolicy)).toBe('certificate-revoked');
	});

	it('returns verified when all checks pass', () => {
		const detail = mockDetail({
			signatureValueStatus: 'verified',
			referenceChecks: [mockReferenceCheck({ digestStatus: 'verified' })],
			certificateTrustStatus: 'trusted',
		});
		expect(computeDetailStatus(detail, defaultPolicy)).toBe('verified');
	});

	it('returns structural-only when no digests are verified and signature is not checked', () => {
		const detail = mockDetail({
			signatureValueStatus: 'not-checked',
			referenceChecks: [mockReferenceCheck({ digestStatus: 'insufficient-data' })],
			certificateTrustStatus: 'not-checked',
		});
		expect(computeDetailStatus(detail, defaultPolicy)).toBe('structural-only');
	});

	it('does not verify when any signed reference is unverified', () => {
		const detail = mockDetail({
			referenceChecks: [
				mockReferenceCheck(),
				mockReferenceCheck({ uri: '#idPackageObject', digestStatus: 'insufficient-data' }),
			],
		});
		expect(computeDetailStatus(detail, defaultPolicy)).toBe('structural-only');
	});

	it('does not verify a signature without references', () => {
		const detail = mockDetail({ referenceChecks: [] });
		expect(computeDetailStatus(detail, defaultPolicy)).toBe('structural-only');
	});

	it('reports a content type mismatch as a digest mismatch', () => {
		const detail = mockDetail({
			referenceChecks: [mockReferenceCheck({ digestStatus: 'content-type-mismatch' })],
		});
		expect(computeDetailStatus(detail, defaultPolicy)).toBe('digest-mismatch');
	});

	it('reports coverage gaps ahead of certificate trust', () => {
		const detail = mockDetail({
			coverageGaps: [{ kind: 'unsigned-part', part: 'ppt/slides/slide1.xml' }],
			certificateTrustStatus: 'untrusted',
		});
		expect(computeDetailStatus(detail, defaultPolicy)).toBe('coverage-incomplete');
	});

	it('returns certificate-untrusted when trust is untrusted but signature is verified', () => {
		const detail = mockDetail({
			signatureValueStatus: 'verified',
			certificateTrustStatus: 'untrusted',
		});
		expect(computeDetailStatus(detail, defaultPolicy)).toBe('certificate-untrusted');
	});

	it('returns timestamp-untrusted when timestampAuthorityStatus is untrusted', () => {
		const detail = mockDetail({
			timestampAuthorityStatus: 'untrusted',
		});
		expect(computeDetailStatus(detail, defaultPolicy)).toBe('timestamp-untrusted');
	});

	it('returns timestamp-invalid when timestampAuthorityStatus is invalid', () => {
		const detail = mockDetail({
			timestampAuthorityStatus: 'invalid',
		});
		expect(computeDetailStatus(detail, defaultPolicy)).toBe('timestamp-invalid');
	});

	it('prioritizes signature-invalid over reference-missing', () => {
		const detail = mockDetail({
			signatureValueStatus: 'invalid',
			missingPartReferences: ['missing.xml'],
		});
		expect(computeDetailStatus(detail, defaultPolicy)).toBe('signature-invalid');
	});

	it('prioritizes reference-missing over digest-mismatch', () => {
		const detail = mockDetail({
			missingPartReferences: ['missing.xml'],
			referenceChecks: [mockReferenceCheck({ digestStatus: 'mismatch' })],
		});
		expect(computeDetailStatus(detail, defaultPolicy)).toBe('reference-missing');
	});

	describe('with requireRevocationCheck policy', () => {
		const strictPolicy: SignatureValidationPolicy = {
			requireRevocationCheck: true,
			failOnRevocationUnknown: true,
			requireTimestamp: false,
		};

		it('returns certificate-untrusted when revocation is unknown and policy requires check', () => {
			const detail = mockDetail({
				certificateRevocationStatus: 'unknown',
			});
			expect(computeDetailStatus(detail, strictPolicy)).toBe('certificate-untrusted');
		});
	});

	describe('with requireTimestamp policy', () => {
		const timestampPolicy: SignatureValidationPolicy = {
			requireRevocationCheck: false,
			failOnRevocationUnknown: false,
			requireTimestamp: true,
		};

		it('returns timestamp-untrusted when timestamp is not-present and policy requires it', () => {
			const detail = mockDetail({
				timestampAuthorityStatus: 'not-present',
			});
			expect(computeDetailStatus(detail, timestampPolicy)).toBe('timestamp-untrusted');
		});
	});
});

describe('computeVerificationStatus', () => {
	it('returns signature-invalid if any detail has that status', () => {
		const details = [
			mockFullDetail({ status: 'verified' }),
			mockFullDetail({ status: 'signature-invalid' }),
		];
		expect(computeVerificationStatus(details)).toBe('signature-invalid');
	});

	it('returns verified-trusted if all details are verified', () => {
		const details = [
			mockFullDetail({ status: 'verified' }),
			mockFullDetail({ status: 'verified' }),
		];
		expect(computeVerificationStatus(details)).toBe('verified-trusted');
	});

	it('returns verified-untrusted if some details are certificate-untrusted', () => {
		const details = [
			mockFullDetail({ status: 'verified' }),
			mockFullDetail({ status: 'certificate-untrusted' }),
		];
		expect(computeVerificationStatus(details)).toBe('verified-untrusted');
	});

	it('does not report verified-untrusted beside an unverifiable signature', () => {
		const details = [
			mockFullDetail({ status: 'certificate-untrusted' }),
			mockFullDetail({ status: 'structural-only' }),
		];
		expect(computeVerificationStatus(details)).toBe('present-not-verified');
	});

	it('returns coverage-incomplete ahead of a verified signature', () => {
		const details = [
			mockFullDetail({ status: 'verified' }),
			mockFullDetail({ status: 'coverage-incomplete' }),
		];
		expect(computeVerificationStatus(details)).toBe('coverage-incomplete');
	});

	it('returns present-not-verified for structural-only details', () => {
		const details = [mockFullDetail({ status: 'structural-only' })];
		expect(computeVerificationStatus(details)).toBe('present-not-verified');
	});

	it('returns certificate-revoked if any detail is revoked', () => {
		const details = [
			mockFullDetail({ status: 'verified' }),
			mockFullDetail({ status: 'certificate-revoked' }),
		];
		expect(computeVerificationStatus(details)).toBe('certificate-revoked');
	});

	it('returns digest-mismatch if any detail has that status', () => {
		const details = [mockFullDetail({ status: 'digest-mismatch' })];
		expect(computeVerificationStatus(details)).toBe('digest-mismatch');
	});

	it('returns reference-missing if any detail has that status', () => {
		const details = [mockFullDetail({ status: 'reference-missing' })];
		expect(computeVerificationStatus(details)).toBe('reference-missing');
	});

	it('returns timestamp-invalid if any detail has that status', () => {
		const details = [mockFullDetail({ status: 'timestamp-invalid' })];
		expect(computeVerificationStatus(details)).toBe('timestamp-invalid');
	});

	it('returns timestamp-untrusted if any detail has that status', () => {
		const details = [mockFullDetail({ status: 'timestamp-untrusted' })];
		expect(computeVerificationStatus(details)).toBe('timestamp-untrusted');
	});

	it('prioritizes signature-invalid over certificate-revoked', () => {
		const details = [
			mockFullDetail({ status: 'signature-invalid' }),
			mockFullDetail({ status: 'certificate-revoked' }),
		];
		expect(computeVerificationStatus(details)).toBe('signature-invalid');
	});

	it('prioritizes certificate-revoked over timestamp-invalid', () => {
		const details = [
			mockFullDetail({ status: 'certificate-revoked' }),
			mockFullDetail({ status: 'timestamp-invalid' }),
		];
		expect(computeVerificationStatus(details)).toBe('certificate-revoked');
	});
});

describe('integrity comes before policy, certificates and timestamps', () => {
	type Fields = Parameters<typeof mockDetail>[0];

	const flags = [false, true];
	const POLICIES: SignatureValidationPolicy[] = flags.flatMap((requireRevocationCheck) =>
		flags.flatMap((failOnRevocationUnknown) =>
			flags.map((requireTimestamp) => ({
				requireRevocationCheck,
				failOnRevocationUnknown,
				requireTimestamp,
			})),
		),
	);
	const TRUST: SignatureDetail['certificateTrustStatus'][] = [
		'trusted',
		'untrusted',
		'not-checked',
	];
	const REVOCATION: CertificateRevocationStatus[] = [
		'good',
		'revoked',
		'unknown',
		'not-checked',
		'error',
	];
	const TIMESTAMP: TimestampAuthorityStatus[] = [
		'valid',
		'invalid',
		'not-present',
		'not-checked',
		'error',
		'untrusted',
	];
	/** Every certificate, revocation, timestamp and policy state a signature can be in. */
	const OUTSIDE_INTEGRITY = POLICIES.flatMap((policy) =>
		TRUST.flatMap((certificateTrustStatus) =>
			REVOCATION.flatMap((certificateRevocationStatus) =>
				TIMESTAMP.map((timestampAuthorityStatus) => ({
					policy,
					fields: { certificateTrustStatus, certificateRevocationStatus, timestampAuthorityStatus },
				})),
			),
		),
	);
	const reference = (digestStatus: SignatureReferenceCheck['digestStatus']) => [
		mockReferenceCheck({ digestStatus }),
	];
	const FAILURES: [string, Fields, SignatureDetailStatus, DigitalSignatureVerificationStatus][] = [
		[
			'an invalid signature value',
			{ signatureValueStatus: 'invalid', referenceChecks: [] },
			'signature-invalid',
			'signature-invalid',
		],
		[
			'a missing part',
			{ missingPartReferences: ['/a.xml'], referenceChecks: reference('missing-part') },
			'reference-missing',
			'reference-missing',
		],
		[
			'a digest mismatch',
			{ referenceChecks: reference('mismatch') },
			'digest-mismatch',
			'digest-mismatch',
		],
		[
			'a content type mismatch',
			{ referenceChecks: reference('content-type-mismatch') },
			'digest-mismatch',
			'digest-mismatch',
		],
		[
			'an unsupported transform',
			{ referenceChecks: reference('unsupported-transform') },
			'structural-only',
			'present-not-verified',
		],
		[
			'an unsupported digest',
			{ referenceChecks: reference('unsupported-algorithm') },
			'structural-only',
			'present-not-verified',
		],
		[
			'a reference without a digest',
			{ referenceChecks: reference('insufficient-data') },
			'structural-only',
			'present-not-verified',
		],
		['no references', { referenceChecks: [] }, 'structural-only', 'present-not-verified'],
		[
			'an unchecked signature value',
			{ signatureValueStatus: 'not-checked', referenceChecks: [] },
			'structural-only',
			'present-not-verified',
		],
		[
			'a coverage gap',
			{ coverageGaps: [{ kind: 'unsigned-part', part: 'ppt/slides/slide2.xml' }] },
			'coverage-incomplete',
			'coverage-incomplete',
		],
	];
	const PASSED_INTEGRITY: SignatureDetailStatus[] = [
		'verified',
		'certificate-untrusted',
		'certificate-revoked',
		'timestamp-invalid',
		'timestamp-untrusted',
	];

	it.each(FAILURES)('keeps %s whatever else holds', (_label, failure, status) => {
		for (const { policy, fields } of OUTSIDE_INTEGRITY) {
			expect(computeDetailStatus(mockDetail({ ...fields, ...failure }), policy)).toBe(status);
		}
	});

	it.each(FAILURES)('reports %s beside any other signature', (_label, _failure, status, report) => {
		expect(computeVerificationStatus([mockFullDetail({ status })])).toBe(report);
		for (const other of PASSED_INTEGRITY) {
			const details = [mockFullDetail({ status }), mockFullDetail({ status: other })];
			expect(computeVerificationStatus(details)).toBe(report);
			expect(computeVerificationStatus(details.reverse())).toBe(report);
		}
	});

	const STRICT: SignatureValidationPolicy = {
		requireRevocationCheck: true,
		failOnRevocationUnknown: true,
		requireTimestamp: true,
	};
	const REVOCATION_ONLY = { ...defaultPolicy, requireRevocationCheck: true };
	const TIMESTAMP_ONLY = { ...defaultPolicy, requireTimestamp: true };

	it.each<[string, Fields, SignatureValidationPolicy, SignatureDetailStatus]>([
		['nothing wrong', {}, STRICT, 'verified'],
		[
			'no policy and nothing checked',
			{ certificateRevocationStatus: 'unknown', timestampAuthorityStatus: 'not-present' },
			defaultPolicy,
			'verified',
		],
		[
			'a revoked signer',
			{ certificateRevocationStatus: 'revoked' },
			defaultPolicy,
			'certificate-revoked',
		],
		[
			'unknown revocation under the policy',
			{ certificateRevocationStatus: 'unknown' },
			REVOCATION_ONLY,
			'certificate-untrusted',
		],
		[
			'a failed revocation check under the policy',
			{ certificateRevocationStatus: 'error' },
			REVOCATION_ONLY,
			'certificate-untrusted',
		],
		[
			'an unchecked signer under the policy',
			{ certificateRevocationStatus: 'not-checked' },
			REVOCATION_ONLY,
			'certificate-untrusted',
		],
		[
			'an invalid timestamp',
			{ timestampAuthorityStatus: 'invalid' },
			defaultPolicy,
			'timestamp-invalid',
		],
		[
			'an untrusted timestamp',
			{ timestampAuthorityStatus: 'untrusted' },
			defaultPolicy,
			'timestamp-untrusted',
		],
		[
			'no timestamp under the policy',
			{ timestampAuthorityStatus: 'not-present' },
			TIMESTAMP_ONLY,
			'timestamp-untrusted',
		],
		[
			'an untrusted signer',
			{ certificateTrustStatus: 'untrusted' },
			defaultPolicy,
			'certificate-untrusted',
		],
		[
			'an unchecked chain',
			{ certificateTrustStatus: 'not-checked' },
			defaultPolicy,
			'structural-only',
		],
	])('gives an intact signature its own status: %s', (_label, fields, policy, status) => {
		expect(computeDetailStatus(mockDetail(fields), policy)).toBe(status);
	});

	it('only reports a verified package when integrity holds', () => {
		const states: [Fields, boolean][] = [
			[{}, true],
			...FAILURES.map(([, failure]): [Fields, boolean] => [failure, false]),
		];
		for (const [integrity, intact] of states) {
			for (const { policy, fields } of OUTSIDE_INTEGRITY) {
				const status = computeDetailStatus(mockDetail({ ...fields, ...integrity }), policy);
				const report = computeVerificationStatus([mockFullDetail({ status })]);
				if (!intact) expect(report).not.toMatch(/^verified-/);
				if (report === 'verified-trusted') expect(fields.certificateTrustStatus).toBe('trusted');
			}
		}
	});
});
