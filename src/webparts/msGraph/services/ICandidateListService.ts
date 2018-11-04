import { Item } from "@pnp/sp";
import { ISPListCandidate, ISPListCandidateListItem } from "./responseTypes/ISPListCandidate";

/**
 * A simple Service that requests information
 * from a List.
 */
export interface ICandidateListService {
    /**
     * Retrives the Author/Candidate from the Author field
     * of the List Item.
     */
    getUserPrincipal(listItemId: number): Promise<ISPListCandidate>;

    getListOfCandidates(): Promise<ISPListCandidateListItem[]>;
}