import React from "react";
import { useAppContext } from "@/context"
import { AuthenticatedTemplate, UnauthenticatedTemplate } from "@azure/msal-react";
import { Stack, Text, PrimaryButton, StackItem } from '@fluentui/react';
import { Directory } from "./Directory";
import { Login } from "./Login";

export const Authenticate = () => {
    const app = useAppContext();
    return <>
        <AuthenticatedTemplate>
            <Directory {...app} />
        </AuthenticatedTemplate>
        <UnauthenticatedTemplate>
            <Login {...app} ></Login>
        </UnauthenticatedTemplate>
    </>
}