/**
 * An interface that is just here to provide TypeScript error checking.
 */
export interface IGraphCandidate {
    businessPhones: string[];
    displayName: string;
    givenName: string;
    id: string;
    mail: string;
    mobilePhone: string;
    officeLocation: string;
    preferrefdLanguage: string;
    surname: string;
    userPrincipalName: string;
}