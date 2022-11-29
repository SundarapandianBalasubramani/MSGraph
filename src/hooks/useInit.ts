import { useAppContext } from "@/context";
import { getUserAndTeamsInformation, UserInfoResponse } from "@/service/user";
import { BatchRequestContent, BatchRequestStep } from "@microsoft/microsoft-graph-client"
import { useEffect, useState } from "react";

const userInfoRequestStep: BatchRequestStep = {
    id: "1",
    request: new Request("/me?$select=displayName,mail,mailboxSettings,userPrincipalName", {
        method: "GET",
        headers: {
            "Content-Type": "application/json"
        }
    })
};

const userAssociatedTeams: BatchRequestStep = {
    id: "2",
    request: new Request("/me/teamwork/associatedTeams?$select=id,displayName,tenantId", {
        method: "GET",
        headers: {
            "Content-Type": "application/json"
        }
    })
};

const listPeople: BatchRequestStep = {
    id: "3",
    request: new Request("/me/people?$select=id,givenName,surname,displayName,jobTitle,department,userPrincipalName&$orderby=displayName desc&$count=true&$top=10", {
        method: "GET",
        headers: {
            "Content-Type": "application/json"
        }
    })
};


let batchRequestContent = new BatchRequestContent([
    userInfoRequestStep,
    userAssociatedTeams,
    listPeople
]);


export const useInit = () => {
    const context = useAppContext();
    const [data, setData] = useState<Partial<UserInfoResponse>>({});
    useEffect(() => {
        const init = async () => {
            const result = await getUserAndTeamsInformation(context.authProvider!, batchRequestContent);            
            setData(result);
        };
        init();
    }, []);
    return data;
}