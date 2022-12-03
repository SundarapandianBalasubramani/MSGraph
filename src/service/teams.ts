import { AuthCodeMSALBrowserAuthenticationProvider } from "@microsoft/microsoft-graph-client/authProviders/authCodeMsalBrowser";
import { Presence, Team } from "microsoft-graph";
import { ensureClient } from "./ensureClient";
import { limit } from "./user";

export async function getUserTeams(authProvider: AuthCodeMSALBrowserAuthenticationProvider): Promise<Team[]> {
    const client = ensureClient(authProvider);

    const teams: Team[] = await client.api('/me/joinedTeams')
        .select('id,displayName,description')
        .get();

    return teams;
}

export type UserIds =
    {
        "ids": string[]
    };

export async function getUsersPresence(authProvider: AuthCodeMSALBrowserAuthenticationProvider, presence: UserIds): Promise<Presence[]> {
    const client = ensureClient(authProvider);
    const usersPresence = await client.api('/communications/getPresencesByUserId')
        .version('beta')
        .post(presence);
    return usersPresence.value;
}

export async function getTeams(authProvider: AuthCodeMSALBrowserAuthenticationProvider, id: string,
    filter: string, skipToken: string): Promise<any> {
    const client = ensureClient(authProvider);
    let people = undefined;
    if (filter.length && skipToken.length)
        people = await client.api(`teams/${id}/members`).filter(filter).skipToken(skipToken).top(limit).get();
    else if (filter.length)
        people = await client.api(`teams/${id}/members`).filter(filter).top(limit).get();
    else if (skipToken.length)
        people = await client.api(`teams/${id}/members`).skipToken(skipToken).top(limit).get();
    else
        people = await client.api(`teams/${id}/members`).top(limit).get();
    return people;
}
