/** Message content passed to the Word editor host after the user confirms delivery. */
export type MailMergeEmailMessage = {
	recipients: readonly string[];
	subject: string;
	attachments: readonly { fileName: string; contentType: string; bytes: Uint8Array }[];
} & ({ text: string; html?: never } | { html: string; text?: never });

/** The host delivers messages through its own service and reports errors by rejecting. */
export interface MailMergeEmailProvider {
	send(message: MailMergeEmailMessage): void | Promise<void>;
}
