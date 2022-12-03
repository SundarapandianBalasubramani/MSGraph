import { BatchRequestContent, BatchRequestStep, BatchResponseContent } from "@microsoft/microsoft-graph-client";
import { AuthCodeMSALBrowserAuthenticationProvider } from "@microsoft/microsoft-graph-client/authProviders/authCodeMsalBrowser";
import { Person, Team, User, } from 'microsoft-graph';
import { ensureClient } from "./ensureClient";

export const limit = 50;

export type UserResponseType = {
    '@odata.context'?: string,
    value: Person[],
    '@odata.nextLink'?: string
}

export type UserInfoResponse = {
    user: User,
    teams: Team[],
    people: UserResponseType,
};

export async function getUser(authProvider: AuthCodeMSALBrowserAuthenticationProvider): Promise<User> {
    const client = ensureClient(authProvider);
    const user: User = await client!.api('/me')
        .select('displayName,mail,mailboxSettings,userPrincipalName')
        .get();

    return user;
}

export async function getUserPhoto(authProvider: AuthCodeMSALBrowserAuthenticationProvider, id: string): Promise<any> {
    const client = ensureClient(authProvider);
    const photo: User = await client!.api(`/users/${id}/photo/$value`)
        .get();
    return photo;
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
        result.people = await peopleResponse.json();
    }
    return result;
}

export async function getOrgPeople(authProvider: AuthCodeMSALBrowserAuthenticationProvider,
    filter: string, skipToken: string): Promise<UserResponseType> {
    const client = ensureClient(authProvider);
    let people;
    if(filter.length && skipToken.length){
         people = await client!.api(`/users`)        
        .select('id,givenName,surname,displayName,jobTitle,department,userPrincipalName')
        .top(limit)
        .filter(`startswith(displayName, \'${filter.toLowerCase()}\')`)
        .skipToken(skipToken)
        .get();
    }
    else if(filter.length){
         people = await client!.api(`/users`)        
        .select('id,givenName,surname,displayName,jobTitle,department,userPrincipalName')
        .top(limit)
        .filter(`startswith(displayName, \'${filter.toLowerCase()}\')`)        
        .get();
    }
    else if(skipToken.length){
        people = await client!.api(`/users`)        
       .select('id,givenName,surname,displayName,jobTitle,department,userPrincipalName')
       .top(limit)       
       .skipToken(skipToken)
       .get();
   }
    else{
         people = await client!.api(`/users`)        
        .select('id,givenName,surname,displayName,jobTitle,department,userPrincipalName')
        .top(limit)               
        .get();
    }
  
    return people;
}

export async function getPeople(authProvider: AuthCodeMSALBrowserAuthenticationProvider,
    filter: string, skipToken: string): Promise<UserResponseType> {
    const client = ensureClient(authProvider);
    let people;
    if (filter.length && skipToken.length) {
        people = await client!.api(`/me/people`)
            .count()
            .select('id,givenName,surname,displayName,jobTitle,department,userPrincipalName')
            .top(limit)
            .search(filter)
            .orderby('displayName').skipToken(skipToken)
            .get();
    }
    else if (filter.length) {
        people = await client!.api(`/me/people`)
            .count()
            .select('id,givenName,surname,displayName,jobTitle,department,userPrincipalName')
            .top(limit)
            .search(filter)
            .orderby('displayName')
            .get();
    }
    else if (skipToken.length) {
        people = await client!.api(`/me/people`)
            .count()
            .select('id,givenName,surname,displayName,jobTitle,department,userPrincipalName')
            .top(limit)            
            .orderby('displayName').skipToken(skipToken)
            .get();
    }
    else  {
        people = await client!.api(`/me/people`)
            .count()
            .select('id,givenName,surname,displayName,jobTitle,department,userPrincipalName')
            .top(limit)         
            .get();
    }   
    return people;
}