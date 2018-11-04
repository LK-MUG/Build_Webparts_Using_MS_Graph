/**
 * An interface that is just here to provide TypeScript error checking.
 */
export interface ISPListCandidate {
    EMail: string;
    Id: number;
    Name: string;
    Title: string;
    // "odata.id": string;
    // "odata.type": string;
}

/**
 * An interface that is just here to provide TypeScript error checking.
 */
export interface ISPListCandidateListItem {
    Author: ISPListCandidate;
    // "Author@odata.navigationLinkUrl": string;
    // "odata.editLink": string;
    // "odata.etag": string;
    // "odata.id": string;
    // "odata.type": string;
}