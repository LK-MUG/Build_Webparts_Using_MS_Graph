import { ICandidateListService } from "./ICandidateListService";
import { SPRest, Item } from "@pnp/sp";
import { ISPListCandidate, ISPListCandidateListItem } from "./responseTypes/ISPListCandidate";

export class CandidateListService implements ICandidateListService {
    private spClient: SPRest;
    private listName: string;

    /**
     * We use dependency injection here, so that the only integration
     * with our API/Service and SharePoint is the constructor and not
     * the whole API.
     * @param spClient 
     */
    public constructor(spClient: SPRest, listName: string) {
        this.spClient = spClient;
        this.listName = listName;
    }

    /**
     * Takes an ID number of the List and returns user principal of the
     * candidate who created that ListItem.
     * @param candidatesListId 
     */
    public getUserPrincipal(listItemId: number): Promise<ISPListCandidate> {
        return new Promise<ISPListCandidate>((resolve: any, reject: any) => {
            // we only want the Authors email/principal from the list item, every thing else is
            // irrelevant.
            this.spClient.web.lists.getByTitle(this.listName).items.getById(listItemId)
            // Author is the field to expand, select Author/Id is the way to retrieve that from the expansion.
            .select("Author/Id,Author/Title,Author/Name,Author/EMail").expand("Author").get()
            .then((item: any | Item) => {
                let userPrincipal: ISPListCandidate = {
                    EMail: item.Author.EMail,
                    Id: item.Author.Id,
                    Name: item.Author.Name,
                    Title: item.Author.Title
                };
                resolve(userPrincipal);
            })
            .catch((error) => {
                console.error(error);
                reject(error);
            });
        });
    }

    public getListOfCandidates(): Promise<ISPListCandidateListItem[]> {
        return new Promise<ISPListCandidateListItem[]>((resolve: any, reject: any) => {
            this.spClient.web.lists.getByTitle(this.listName).items
            .select("Author/Id,Author/Title,Author/Name,Author/EMail").expand("Author").get()
            .then((candidates: any[]) => {
                let tempCandidates: ISPListCandidateListItem[] = candidates.map((currentCandidate) => {
                    let author: ISPListCandidate = {
                        EMail: currentCandidate.Author.EMail,
                        Id: currentCandidate.Author.Id,
                        Name: currentCandidate.Author.Name,
                        Title: currentCandidate.Author.Title
                    };

                    return {
                        Author: author
                    };
                });
                resolve(tempCandidates);
            })
            .catch((error) => {
                console.error(error);
                reject(error);
            });
        });
    }
}