// Compatibility exports: all document operations are owned by OOXML core.
import { apiToolsCustomShowTools as operations } from 'ooxml-core/pptx/automation';

export type ManageCustomShowsParams = operations.ManageCustomShowsParams;
export type CustomShowInfo = operations.CustomShowInfo;
export type ManageCustomShowsResult = operations.ManageCustomShowsResult;
export const manageCustomShows = operations.manageCustomShows;
