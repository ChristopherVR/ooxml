import type { Comment } from '../model';

/** A pasted copy retains metadata but must not take the source thread's identities. */
export function resetCommentIdentities(comment: Comment): void {
	if (comment.replies) comment.regenerateThreadIds = true;
}
