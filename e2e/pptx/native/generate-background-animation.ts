import { backgroundAnimationDeck } from '../support/background-animation-deck';

const destination = process.argv[2];
if (!destination) throw new Error('Usage: bun generate-background-animation.ts output.pptx');
await Bun.write(destination, await backgroundAnimationDeck());
