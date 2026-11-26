import { readRecipientFile } from './recipient-data';

self.addEventListener(
	'message',
	(event: MessageEvent<File>) => {
		// Only the page that created this worker may send it files.
		if (event.origin && event.origin !== self.location.origin) return;
		void readRecipientFile(event.data).then(
			(data) => self.postMessage(data),
			() => self.postMessage(null),
		);
	},
	{ once: true },
);
