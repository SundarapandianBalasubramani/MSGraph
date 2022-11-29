import { AppContext } from "@/context"
import { useInit } from "@/hooks/useInit"
import { getTeamsPersons, getUsersPresence } from "@/service/teams"
import { getPeople } from "@/service/user"
import { Stack, StackItem, Text, PrimaryButton, DefaultButton, Toggle } from "@fluentui/react"
import { Person, Presence } from "microsoft-graph"
import { useEffect, useState } from "react"
import { Departments } from "./Department"
import { Loader } from "./Loader"
import { Search } from "./Search"
import { User } from "./User"

export type PageSet = {
    [key: number]: string | undefined
}
//$filter=startswith(displayName,'a')
const peopleUrl = "$select=id,givenName,surname,displayName,jobTitle,department,userPrincipalName&$orderby=displayName desc&$count=true&$top=100";

export const Directory = (props: AppContext) => {
    const data = useInit();
    const [showLoader, setShowLoader] = useState(true);
    const [persons, setPerson] = useState<Person[]>([]);
    const [selectedTeam, setTeam] = useState('');
    const [filter, setFilter] = useState('');
    const [letter, setLetter] = useState('');
    const [next, setNext] = useState<PageSet>({});
    const [page, setPage] = useState(1);
    const [userSource, setUserSource] = useState(true);
    const [url, setUrl] = useState('/me/people');
    useEffect(() => {
        if (data && data.people) {
            setNext({ 2: data.next });
            setPerson(data.people);
        }
        setShowLoader(false);
    }, [data.user]);

    useEffect(() => {
        getPeopleData(url);

    }, [userSource]);

    useEffect(() => {
        if (filter.length > 0) {
            if (selectedTeam.length === 0)
                getPeopleData(url + '?' + filter + peopleUrl);
            else
                getTeamsPeopleData(selectedTeam);
        }

    }, [filter, url]);

    useEffect(() => {
        getTeamsPeopleData(selectedTeam);
    }, [selectedTeam]);

    const getPeopleData = async (url: string) => {
        const result = await getPeople(props.authProvider!, url);
        console.log(result);
        const pageNo: number = page + 1;
        if (result['@odata.nextLink']) {
            setNext({ ...next, [pageNo]: result['@odata.nextLink'] });
            console.log({ ...next, [pageNo]: result['@odata.nextLink'] });
        }
        setPerson(result.value);
        setShowLoader(false);
    }

    const getTeamsPeopleData = async (team: string) => {
        const result = await getTeamsPersons(props.authProvider!, team, letter);
        console.log(result);
        const pageNo: number = page + 1;
        if (result['@odata.nextLink']) {
            setNext({ ...next, [pageNo]: result['@odata.nextLink'] });
            console.log({ ...next, [pageNo]: result['@odata.nextLink'] });
        }
        setPerson(result.value);
        setShowLoader(false);
    }



    const givenNameStartsWith = (l: string) => {
        setLetter(l);
        if (!userSource)
            setFilter(`$filter=startswith(givenName, \'${l.toLowerCase()}\') or startswith(surName, \'${l.toLowerCase()}\')&ConsistencyLevel=eventual&`);
        else
            setFilter(`$search=${l.toLowerCase()}&`);
        setPage(1);
        setNext({});
        setPerson([]);
        setShowLoader(true);
    }

    const teams = (team: string) => {
        setNext({});
        setPerson([]);
        setFilter('');
        setLetter('');
        setTeam(team);
        setShowLoader(true);
    }


    const nextClick = () => {
        setPage(page + 1);
        setShowLoader(true);
        getPeopleData(next[(page + 1)]!);
    }

    const previousClick = () => {
        setShowLoader(true);
        if (page > 2) {
            setPage(page - 1);
            getPeopleData(next[(page - 1)]!);
        } else {
            setPage(1);
            setNext({});
            getPeopleData(url + '?' + filter + peopleUrl);
        }
    }


    const onToggle = (ev: React.MouseEvent<HTMLElement>, checked?: boolean) => {
        setLetter('');
        setFilter('');
        setPage(1);
        setTeam('');
        setNext({});
        setPerson([]);
        setShowLoader(true);
        setUserSource(checked as boolean);
        setUrl(checked ? "/me/people" : "/users");
    }

    const reset = () => {
        setLetter('');
        setFilter('');
        setPage(1);
        setTeam('');
        setUserSource(true);
        setNext({});
        setPerson([]);
        setShowLoader(true);
        getPeopleData(url + '?' + filter + peopleUrl);
    }

    const search = (k: string) => {
        if (k.length > 2)
            givenNameStartsWith(k);
    }

    const onChange = (event?: React.ChangeEvent<HTMLInputElement>, newValue?: string) => {

    }

    return <Stack >
        <StackItem align="center">
            <Text block variant={"xLargePlus"}>Graph API User Directory App</Text>
        </StackItem>
        <StackItem align="end" style={{ marginBottom: 20 }}>
            <Text variant="large">Welcome {data?.user?.displayName || ''}!</Text>
        </StackItem>
        <StackItem align="end" style={{ marginBottom: 20 }}>
            <PrimaryButton onClick={props.signOut!}>Sign out</PrimaryButton>
        </StackItem>
        <StackItem  >
            <Stack horizontal tokens={{ childrenGap: 20, padding: 20 }}>
                <StackItem grow={3} >
                    <Stack tokens={{ childrenGap: 20 }}>
                        <Text style={{ width: 200 }} variant={"medium"}>Organization and Related People works only if no team is selected </Text>
                        <Toggle checked={userSource} offText="Organization" onText="Frequent Contacts" onChange={onToggle} />
                        <Departments data={data.teams} teams={teams} team={selectedTeam} />
                        <PrimaryButton style={{ width: 250 }} onClick={reset}>Reset</PrimaryButton>
                    </Stack>
                </StackItem>
                <StackItem grow={4} >
                    <Search onChange={onChange} onClick={givenNameStartsWith} letter={letter} searchValue={letter} onSearch={search} />
                    <Stack wrap tokens={{ childrenGap: 20, padding: 40 }} horizontal >
                        {persons.map(p => <StackItem key={p.id}><User key={p.id} user={p} authProvider={props.authProvider} /></StackItem>)}
                        {persons.length === 0 && !showLoader && <h2>No Results found</h2>}
                    </Stack>
                </StackItem>
            </Stack>
        </StackItem>
        <StackItem align="center">
            <Stack tokens={{ childrenGap: 20 }} horizontal>
                {/* <DefaultButton onClick={previousClick} disabled={page === 1}>Previous</DefaultButton>
                <Text block variant={"xLargePlus"}>{persons.length} Users of Page {page}</Text>
                <DefaultButton onClick={nextClick} disabled={!Boolean(next[(page + 1)])}>Next</DefaultButton> */}
            </Stack>
        </StackItem>
        <Loader show={showLoader} />
    </Stack>
}