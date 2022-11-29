import { Client } from "@microsoft/microsoft-graph-client";
import { AuthCodeMSALBrowserAuthenticationProvider } from "@microsoft/microsoft-graph-client/authProviders/authCodeMsalBrowser";

let graphClient: Client | undefined = undefined;
export function ensureClient(authProvider: AuthCodeMSALBrowserAuthenticationProvider):Client {
    if (!graphClient) {
        graphClient = Client.initWithMiddleware({
            authProvider: authProvider
        });
    }

    return graphClient;
}
