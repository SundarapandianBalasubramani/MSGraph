import { createContext, MouseEventHandler, useContext, useEffect, useState } from "react";
import { AuthCodeMSALBrowserAuthenticationProvider } from '@microsoft/microsoft-graph-client/authProviders/authCodeMsalBrowser';
import { useMsal } from "@azure/msal-react";
import { InteractionType, PublicClientApplication } from "@azure/msal-browser";
import { config } from "@/auth";
import { getUser, UserInfoResponse } from "@/service/user";
import { User } from 'microsoft-graph';
// <AppContextSnippet>
export interface AppUser {
    displayName?: string,
    email?: string,
    avatar?: string,
    timeZone?: string,
    timeFormat?: string
};

export interface AppError {
    message: string,
    debug?: string
};

export type AppContext = {
    error?: AppError;
    signIn?: MouseEventHandler<HTMLElement>;
    signOut?: MouseEventHandler<HTMLElement>;
    displayError?: Function;
    clearError?: Function;
    authProvider?: AuthCodeMSALBrowserAuthenticationProvider;
    data?: Partial<UserInfoResponse>
}

const appContext = createContext<AppContext>({
    error: undefined,
    signIn: undefined,
    signOut: undefined,
    displayError: undefined,
    clearError: undefined,
    authProvider: undefined
});

export function useAppContext(): AppContext {
    return useContext(appContext);
}

interface ProvideAppContextProps {
    children: React.ReactNode;
}

export default function ProvideAppContext({ children }: ProvideAppContextProps) {
    const auth = useProvideAppContext();
    return (
        <appContext.Provider value={auth} >
            {children}
        </appContext.Provider>
    );
}
// </AppContextSnippet>

function useProvideAppContext() {
    const msal = useMsal();
    //const [user, setUser] = useState<AppUser | undefined>(undefined);
    const [error, setError] = useState<AppError | undefined>(undefined);

    const displayError = (message: string, debug?: string) => {
        setError({ message, debug });
    }

    const clearError = () => {
        setError(undefined);
    }

    // <AuthProviderSnippet>
    // Used by the Graph SDK to authenticate API calls
    const authProvider = new AuthCodeMSALBrowserAuthenticationProvider(
        msal.instance as PublicClientApplication,
        {
            account: msal.instance.getActiveAccount()!,
            scopes: config.scopes,
            interactionType: InteractionType.Popup
        }
    );
    // </AuthProviderSnippet>

    // <UseEffectSnippet>
    // useEffect(() => {
    //     const checkUser = async () => {
    //         if (!user) {
    //             try {
    //                 // Check if user is already signed in
    //                 const account = msal.instance.getActiveAccount();
    //                 if (account) {
    //                     // Get the user from Microsoft Graph
    //                     const user: User = await getUser(authProvider);

    //                     setUser({
    //                         displayName: user.displayName || '',
    //                         email: user.mail || user.userPrincipalName || '',
    //                         timeFormat: user.mailboxSettings?.timeFormat || 'h:mm a',
    //                         timeZone: user.mailboxSettings?.timeZone || 'UTC'
    //                     });
    //                 }
    //             } catch (err: any) {
    //                 displayError(err.message);
    //             }
    //         }
    //     };
    //     checkUser();
    // }, []);
    // </UseEffectSnippet>

    // <SignInSnippet>
    const signIn = async () => {
        await msal.instance.loginPopup({
            scopes: config.scopes,
            prompt: 'select_account'
        });

        // Get the user from Microsoft Graph
        // const user: User = await getUser(authProvider);

        // setUser({
        //     displayName: user.displayName || '',
        //     email: user.mail || user.userPrincipalName || '',
        //     timeFormat: user.mailboxSettings?.timeFormat || '',
        //     timeZone: user.mailboxSettings?.timeZone || 'UTC'
        // });
    };
    // </SignInSnippet>

    // <SignOutSnippet>
    const signOut = async () => {
        await msal.instance.logoutPopup();
        // setUser(undefined);
    };
    // </SignOutSnippet>

    return {
        error,
        signIn,
        signOut,
        displayError,
        clearError,
        authProvider
    };
}