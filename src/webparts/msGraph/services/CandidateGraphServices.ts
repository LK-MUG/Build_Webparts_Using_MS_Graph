import { ICandidateGraphService } from "./ICandidateGraphServices";
import { GraphRest } from "@pnp/graph";
import { IGraphCandidate } from "./responseTypes/IGraphCandidate";

/**
 * This class Reads the internal inforamtion about a candidate 
 * from using the MS-Graph REST API.
 */
export class CandidateGraphService implements ICandidateGraphService {
    private restClient: GraphRest;

    /**
     * The constructor requires an instantiated GraphRest client,
     * because this class does not deal with getting the O-Auth token.
     * To do the OAuth would mean this API/Service gets coupled with
     * MS-Graph V1. We want to allow MS to change while not changing our
     * code much.
     * @param restClient 
     */
    public constructor (restClient: GraphRest) {
        this.restClient = restClient;
    }

    /**
     * A method to find a user in MS-Grap using the principal/id
     * id sample: b8693aee-48d2-4a2b-a15f-84bf55a8cef0
     * principal sample: someone@lk-mug.onmicrosoft.com
     * @param principal 
     */
    public findUser(principal: string): Promise<IGraphCandidate> {
        return new Promise<IGraphCandidate>((resolve: any, reject: any) => {
            this.restClient.users.concat(`/${principal}`).get()
            .then((user: any) => {
                resolve(user);
            })
            .catch((error: any) => {
                console.error(error);
                reject(error);
            });        
        });
    }

    /**
     * A method to find the candidates manager.
     * @param principal 
     */
    public findManager(principal: string): Promise<any> {
        return new Promise<IGraphCandidate>((resolve: any, reject: any) => {
            this.restClient.users.concat(`/${principal}/manager`).get()
            .then((manager: any) => {
                resolve(manager);
            })
            .catch((error) => {
                console.error(error);
                reject(error);
            });
        });
    }

    /**
     * A method to find the Candidates related people,
     * this can be team-mates for instance.
     * @param principal 
     */
    public findRelatedPeople(principal: string): Promise<IGraphCandidate[]> {
        return new Promise<IGraphCandidate[]>((resolve: any, reject: any) => {
            this.restClient.users.concat(`/${principal}/people`).get()
            .then((people: any[]) => {
                resolve(people);
            })
            .catch((error) => {
                reject(error);
            });
        });
    }

    /**
     * A method to find the Candidates groups.
     * @param principal 
     */
    public findMemberships(principal: string): Promise<any> {
        return new Promise<any>((resolve: any, reject: any) => {
            this.restClient.users.concat(`/${principal}/memberOf`).get()
            .then((memberShips: any[]) => {
                resolve(memberShips);
            })
            .catch((error) => {
                reject(error);
            });
        });
    }

    /**
     * A method to find the Candidates direct reports.
     * The people in the organization that report to the candidate.
     * @param principal 
     */
    public findDirectReports(principal: string): Promise<IGraphCandidate[]> {
        return new Promise<IGraphCandidate[]>((resolve: any, reject: any) => {
            this.restClient.users.concat(`/${principal}/directReports`).get()
            .then((memberShips: any[]) => {
                resolve(memberShips);
            })
            .catch((error) => {
                reject(error);
            });
        });
    }
}