import { FetchOptions } from "@pnp/common";

/**
 * A simple Graph rest service that handles the Parsing options
 * and Authentication for you.
 */
export interface IMSGraphRestService {

    /**
     * A simple fetch that delegates to the normal service
     * but implements some of the options for you.
     * @param url 
     * @param options 
     */
    fetch(url: string, options: FetchOptions): any;
}