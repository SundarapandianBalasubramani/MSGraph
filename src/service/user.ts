import { BatchRequestContent, BatchRequestStep, BatchResponseContent } from "@microsoft/microsoft-graph-client";
import { AuthCodeMSALBrowserAuthenticationProvider } from "@microsoft/microsoft-graph-client/authProviders/authCodeMsalBrowser";
import { Person, Team, User } from 'microsoft-graph';
import { ensureClient } from "./ensureClient";

export type UserInfoResponse = {
    user: User,
    teams: Team[],
    people: Person[],
    next?: string,
};

export async function getUser(authProvider: AuthCodeMSALBrowserAuthenticationProvider): Promise<User> {
    const client = ensureClient(authProvider);

    // Return the /me API endpoint result as a User object
    const user: User = await client!.api('/me')
        // Only retrieve the specific fields needed
        .select('displayName,mail,mailboxSettings,userPrincipalName')
        .get();

    return user;
}

export async function getUserPhoto(authProvider: AuthCodeMSALBrowserAuthenticationProvider, id: string): Promise<any> {
    const client = ensureClient(authProvider);
    // Return the /me API endpoint result as a User object
    const photo: User = await client!.api(`/users/${id}/photo/$value`)
        .get();

    return photo;
}

export async function getUserPhotoAndPresence(authProvider: AuthCodeMSALBrowserAuthenticationProvider, id: string): Promise<any> {
    const client = ensureClient(authProvider);    
    const photo = client!.api(`/users/${id}/photo/$value`)
        .get();
    const presence = client!.api(`/users/${id}/presence`)
        .get();
    const responses = await Promise.all([photo, presence]);
    return responses;
}


export async function getUserAndTeamsInformation(authProvider: AuthCodeMSALBrowserAuthenticationProvider, batchRequestContent: BatchRequestContent): Promise<Partial<UserInfoResponse>> {
    const result: Partial<UserInfoResponse> = {};
    const client = ensureClient(authProvider);
    let content = await batchRequestContent.getContent();
    const batchResponse = await client
        .api('/$batch')
        .post(content)
    let batchResponseContent = new BatchResponseContent(batchResponse);
    let userResponse = batchResponseContent.getResponseById("1");
    let teamsResponse = batchResponseContent.getResponseById("2");
    let peopleResponse = batchResponseContent.getResponseById("3");
    if (userResponse.ok) {
        result.user = await userResponse.json();
    }
    if (teamsResponse.ok) {
        let teams = await teamsResponse.json();
        result.teams = teams.value;
    }
    if (peopleResponse.ok) {
        let people = await peopleResponse.json();
        if (people['@odata.nextLink']) {
            result.next = people['@odata.nextLink'];
        }
        result.people = people.value;
    }
    return result;
}


export async function getPeople(authProvider: AuthCodeMSALBrowserAuthenticationProvider, url: string): Promise<any> {
    const client = ensureClient(authProvider);    
    const people = await client!.api(url)
        .get();

    return people;
}