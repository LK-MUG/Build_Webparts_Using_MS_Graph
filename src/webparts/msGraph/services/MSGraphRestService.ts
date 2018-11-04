import { BearerTokenFetchClient, FetchOptions } from "@pnp/common";
import { AadTokenProvider } from "@microsoft/sp-http";
import { IMSGraphRestService } from "./IMSGraphRestService";

/**
 * A class that makes calling the MS_Graph REST API easier, by dealing with 
 * the the unnecessary complexities for you.
 */
export class MSGraphRestService extends BearerTokenFetchClient implements IMSGraphRestService {

    /**
     * Requires the AadTokenProvider to handle the authentication for you,
     * the rest of the API stuff is handled by BearerTokenFetchClient
     * @param tokenProvider 
     */
    public constructor(private tokenProvider: AadTokenProvider) {
        super(null);
    }

    /**
     * Delegates to the BearerTokenFetchClient after authenticating you.
     * Makes a MS_Graph Fetch HTTP Request.
     * @param url 
     * @param options 
     */
    public fetch(url: string, options: FetchOptions = {}): Promise<Response> {
        return this.tokenProvider.getToken('https://graph.microsoft.com')
            .then((accessToken: string) => {
                this.token = accessToken;
                return super.fetch(url, options);
            });
    }
}