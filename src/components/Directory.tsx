import { AppContext } from "@/context"
import { useInit } from "@/hooks/useInit"
import { getTeams, getUsersPresence } from "@/service/teams"
import { getOrgPeople, getPeople, UserResponseType } from "@/service/user"
import { Stack, StackItem, Text, PrimaryButton, DefaultButton } from "@fluentui/react"
import { Person, Presence } from "microsoft-graph"
import { useEffect, useState } from "react"
import { Filter } from "./Filter"
import { Loader } from "./Loader"
import { Search } from "./Search"
import { User } from "./User"

export type PageSet = {
    [key: number]: string | undefined
}

export type PersonWithPresence = Person & Presence;


export const App = (props: AppContext) => {
    const data = useInit();
    return data ? <Directory {...props} data={data} /> : <></>;
}

export const getSearchType = (): boolean => {
    const type = window.sessionStorage.getItem('type');
    return type === null ? true : false;
}

export const Directory = (props: AppContext) => {
    const [showLoader, setShowLoader] = useState(true);
    const [personsWithPresence, setPersonWithPresence] = useState<PersonWithPresence[]>([]);
    const [selectedTeam, setTeam] = useState('');
    //const [filter, setFilter] = useState('');
    const [letter, setLetter] = useState('');
    // const [next, setNext] = useState<PageSet>({});
    const [page, setPage] = useState(1);

    const updatePresence = async (response: UserResponseType, pageno: number) => {
        if (response && response.value.length) {
            let ids: string[] = selectedTeam.length === 0 ? response.value.map(p => p.id!) : response.value.map((p: any) => p.userId!);
            const usersPresence = await getUsersPresence(props.authProvider!, {
                "ids": ids
            });
            let usersWithPresence: PersonWithPresence[] = [];
            usersWithPresence = response.value.map(p => {
                const userPresence = usersPresence.find(pre => pre.id === p.id);
                return userPresence ? { ...userPresence, ...p } : p;
            });
            setPersonWithPresence(usersWithPresence);
            console.log(usersWithPresence);
            //setSkipTokenValue(response['@odata.nextLink'], pageno);
        }
        setShowLoader(false);
    };


    useEffect(() => {
        window.sessionStorage.removeItem('type');
    }, []);

    useEffect(() => {
        if (props.data && props.data.people) {
            updatePresence(props.data.people, page);
        }
    }, [props.data?.people]);

    useEffect(() => {
        letter.length && fetchData();
    }, [letter]);

    const fetchData = () => {
        if (selectedTeam.length === 0)
            getData(page);
        else
            getTeamsPeopleData(selectedTeam, page);
    }



    // const setSkipTokenValue = (nextLink: string | undefined, pageno: number) => {
    //     let skip = undefined;
    //     if (nextLink)
    //         skip = nextLink.substring(nextLink.lastIndexOf('skipToken=') + 10);
    //     setNext({ ...next, [pageno]: skip });
    //     setPage(pageno);
    //     console.log(pageno, { ...next, [pageno]: skip });
    // }



    // const getSkipTokenValue = (pageno: number): string => {
    //     const val = pageno in next ? next[pageno] : '';
    //     return val ? val : '';
    // }


    const getData = async (pageno: number) => {
        // const result = getSearchType() ? await getPeople(props.authProvider!, letter, getSkipTokenValue(pageno))
        //     : await getOrgPeople(props.authProvider!, filter, getSkipTokenValue(pageno));
        // updatePresence(result, pageno);
        const result = getSearchType() ? await getPeople(props.authProvider!, letter, '')
            : await getOrgPeople(props.authProvider!, letter, '');
        updatePresence(result, pageno);

    }

    const getTeamsPeopleData = async (team: string, pageno: number) => {
        const result = await getTeams(props.authProvider!, team, letter, '');
        console.log(result);
        updatePresence(result, pageno);
    }

    const givenNameStartsWith = (l: string) => {
        setLetter(l);
        setPage(1);
        //setNext({});
        setShowLoader(true);
        //fetchData();
    }

    const teams = (team: string) => {
        setTeam(team);
        //setNext({});        
        setLetter('');
        setShowLoader(true);
        getTeamsPeopleData(team, 1);
    }

    const getNavClick = (pageno: number) => {
        setShowLoader(true);
        if (selectedTeam.length === 0)
            getData(pageno);
        else
            getTeamsPeopleData(selectedTeam, pageno);
    }

    const goNext = async () => {
        getNavClick(page + 1);

    };
    const goPrevious = () => {
        setShowLoader(true);
        if (page > 2) {
            getNavClick(page - 1);
        } else {
            getNavClick(1);
            //setNext({});
        }
    }


    const onToggle = (checked: boolean) => {
        if (checked) window.sessionStorage.setItem('type', "0"); else
            window.sessionStorage.removeItem('type')
        setShowLoader(true);
        setLetter('');
        setPage(1);
        setTeam('');
        //setNext({});
        getData(1);
    }

    const reset = () => {
        setLetter('');
        setPage(1);
        setTeam('');
        //setNext({});
        setShowLoader(true);
    }

    const search = (k: string) => {
        if (k.length !== 0)
            givenNameStartsWith(k);
    }

    const onChange = (event?: React.ChangeEvent<HTMLInputElement>, newValue?: string) => {

    }

    return <>
        <Stack className="container">
            <StackItem align="center">
                <Text block variant={"xLargePlus"}>Graph API User Directory App</Text>
            </StackItem>
            <StackItem align="end" style={{ marginBottom: 20 }}>
                <Text variant="large">Welcome {props.data?.user?.displayName || ''}!</Text>
            </StackItem>
            <StackItem align="end" style={{ marginBottom: 20 }}>
                <PrimaryButton onClick={props.signOut!}>Sign out</PrimaryButton>
            </StackItem>
            <StackItem align="center">
                <Stack horizontal tokens={{ childrenGap: 20 }} style={{ minHeight: 600, paddingTop: 40, paddingBottom: 40 }}>
                    <StackItem >
                        <Search onChange={onChange} onClick={givenNameStartsWith} letter={letter} searchValue={letter} onSearch={search} />
                        <Stack wrap tokens={{ childrenGap: 20, padding: 40 }} horizontal >
                            {personsWithPresence.map(p => <StackItem key={p.id}><User key={p.id} user={p} authProvider={props.authProvider} /></StackItem>)}
                        </Stack>
                        <Stack verticalAlign="center" horizontalAlign="center">
                            {personsWithPresence.length === 0 && !showLoader && <h2>No Results found</h2>}
                        </Stack>
                        {/* {!showLoader &&
                            <Stack tokens={{ childrenGap: 20 }} horizontal verticalAlign="center" horizontalAlign="center" style={{ marginTop: 40 }}>
                                <DefaultButton onClick={goPrevious} disabled={page === 1}>{"<<"}</DefaultButton>
                                <Text block variant={"xLarge"}>Page {page}</Text>
                                <DefaultButton onClick={goNext} disabled={!Boolean(next[(page)])}>{">>"}</DefaultButton>
                            </Stack>
                        } */}
                    </StackItem>
                </Stack>
            </StackItem>
            <Loader show={showLoader} />
        </Stack>
        <Filter onToggle={onToggle} reset={reset} team={selectedTeam} teams={teams} data={props.data?.teams} />
    </>
}