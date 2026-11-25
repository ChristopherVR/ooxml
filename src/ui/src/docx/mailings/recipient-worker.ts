import { readRecipientFile } from './recipient-data';

self.addEventListener(
	'message',
	(event: MessageEvent<File>) => {
		void readRecipientFile(event.data).then(
			(data) => self.postMessage(data),
			() => self.postMessage(null),
		);
	},
	{ once: true },
);
