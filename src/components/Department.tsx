import { INavLink, INavLinkGroup, INavStyles, IRenderGroupHeaderProps, Nav } from "@fluentui/react";
import { Team } from "microsoft-graph";
import { useEffect, useState } from "react";

export interface IDepartment {
    data?: Team[],
    teams(team: string): void,
    team: string
}
const navStyles: Partial<INavStyles> = {
    root: {
        width: 208,
        border: '1px solid #eee',
        padding: 20
    },
};



const onRenderGroupHeader = (group: INavLinkGroup): JSX.Element => {
    return <h3>{group.name}</h3>;
}

export const Departments = (props: IDepartment) => {
    const [links, setLinks] = useState<INavLinkGroup[]>([]);
    useEffect(() => {
        if (props.data) {
            let teams: INavLink[] = props.data.map(t => ({ name: t.displayName, key: t.id } as any));
            setLinks([{ links: teams, name: 'My Teams', }]);
        }
    }, [props.data]);
    const onLinkClick = (ev?: React.MouseEvent<HTMLElement>, item?: INavLink) => {
        if (item) {
            props.teams(item.key as string);
            console.log(item);
        }
    }
    return (
        <Nav
            onRenderGroupHeader={onRenderGroupHeader as any}
            onLinkClick={onLinkClick}
            selectedKey={props.team}
            styles={navStyles}
            groups={links}
        />
    );
}