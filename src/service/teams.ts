import { AuthCodeMSALBrowserAuthenticationProvider } from "@microsoft/microsoft-graph-client/authProviders/authCodeMsalBrowser";
import { Presence, Team } from "microsoft-graph";
import { ensureClient } from "./ensureClient";

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

export async function getTeamsPersons(authProvider: AuthCodeMSALBrowserAuthenticationProvider, id: string, letter: string): Promise<any> {
    const client = ensureClient(authProvider);
    const people = letter.length > 0 ? await client.api(`teams/${id}/members`).filter(`startswith(displayName, '${letter}')`)
        .get() : await client.api(`teams/${id}/members`).get();
    return people;
}
