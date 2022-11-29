import { AppContext } from "@/context";
import { PrimaryButton, Stack, StackItem, Text } from "@fluentui/react";

export const Login = (props: AppContext) => {
    return <Stack tokens={{ childrenGap: 40 }} >
        <StackItem align="center">
            <Text block variant={"xLargePlus"}>Graph API User Directory App</Text>
        </StackItem>
        <StackItem align="center" >
            <PrimaryButton onClick={props.signIn!}>Sign in</PrimaryButton>
        </StackItem>
    </Stack>
}