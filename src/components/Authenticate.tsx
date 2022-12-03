import React from "react";
import { useAppContext } from "@/context"
import { AuthenticatedTemplate, UnauthenticatedTemplate } from "@azure/msal-react";
import { App } from "./Directory";
import { Login } from "./Login";

export const Authenticate = () => {
    const app = useAppContext();
    return <>
        <AuthenticatedTemplate>
            <App {...app} />
        </AuthenticatedTemplate>
        <UnauthenticatedTemplate>
            <Login {...app} ></Login>
        </UnauthenticatedTemplate>
    </>
}