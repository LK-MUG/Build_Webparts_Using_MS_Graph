import { IGraphCandidate } from "./responseTypes/IGraphCandidate";

/**
 * A service that makes calling the MS_Graph
 * API simpler.
 */
export interface ICandidateGraphService {
    /**
     * Retrieves the users info from MS_Graph
     * @param principal 
     */
    findUser(principal: string): Promise<IGraphCandidate>;

    /**
     * Retrieves the users managers info from MS_Graph
     * @param principal 
     */
    findManager(principal: string): Promise<any[]>;

    /**
     * Retrieves the users related people info from MS_Graph
     * @param principal 
     */
    findRelatedPeople(principal: string): Promise<IGraphCandidate[]>;

    /**
     * Retrieves the users membership information from MS_Graph
     * @param principal 
     */
    findMemberships(principal: string): Promise<any>;

    /**
     * Retrieves the users direct reports information from MS_Graph
     * @param principal 
     */
    findDirectReports(principal: string): Promise<IGraphCandidate[]>;
}