import { IGraphCandidate } from "../services/responseTypes/IGraphCandidate";
import { ISPListCandidate } from "../services/responseTypes/ISPListCandidate";

/**
 * The properties that cause the CandidateInternalInfo to
 * re-render and change what is displayed
 */
export interface ICandidateInternalInfoState {
    users: ISPListCandidate[];
    candidate: IGraphCandidate;
    selectedId: number;
}
