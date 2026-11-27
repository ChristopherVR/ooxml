/**
 * The host's signing key. The library never reads a certificate store or a private key itself: a
 * host supplies the certificate and a function that signs bytes (Web Crypto with a key the user
 * imported, a hardware token, a signing server). Everything else in the signature is built here.
 */
export interface PackageSigner {
	/** The signer's X.509 certificate, DER encoded; an RSA key of 2048 bits or more is expected. */
	certificate(): Uint8Array | Promise<Uint8Array>;
	/**
	 * An RSASSA-PKCS1-v1_5 signature with SHA-256 (`rsa-sha256` in XML-DSig) over `data`, the
	 * canonical `SignedInfo`. Reject to cancel signing.
	 */
	sign(data: Uint8Array): Promise<Uint8Array>;
	/** Shown in the Sign dialog: whose certificate this is (`Signing as ...`). */
	readonly label?: string;
}

/**
 * A {@link PackageSigner} over a Web Crypto RSASSA-PKCS1-v1_5 SHA-256 private key and its
 * certificate, for hosts that import the user's key (from a PKCS #8 file the user chooses, say).
 */
export function webCryptoPackageSigner(
	privateKey: CryptoKey,
	certificateDer: Uint8Array,
	label?: string,
): PackageSigner {
	return {
		certificate: () => certificateDer,
		async sign(data) {
			const buffer = new ArrayBuffer(data.byteLength);
			new Uint8Array(buffer).set(data);
			return new Uint8Array(
				await globalThis.crypto.subtle.sign({ name: 'RSASSA-PKCS1-v1_5' }, privateKey, buffer),
			);
		},
		...(label ? { label } : {}),
	};
}
