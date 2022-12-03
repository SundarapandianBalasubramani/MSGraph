import { Panel, PrimaryButton, Toggle, Text, IChoiceGroupOption, ChoiceGroup } from "@fluentui/react";
import { useBoolean } from '@fluentui/react-hooks';
import { Team } from "microsoft-graph";
import { Departments } from "./Department";
export interface IFilter {
    data?: Team[];
    teams(team: string): void;
    team: string;
    reset(): void;
    onToggle(checked: boolean): void;
    source?: string
}

const options: IChoiceGroupOption[] = [
    { key: '1', text: 'People Related' },
    { key: '2', text: 'Organization' }
];


export const Filter = ({ data, teams, team, reset, onToggle }: IFilter) => {
    const [isOpen, { setTrue: openPanel, setFalse: dismissPanel }] = useBoolean(false);
    const onChange = (_ev?: React.FormEvent<HTMLElement | HTMLInputElement>, option?: IChoiceGroupOption): void => {
        if (option)
            onToggle(option.key === "2")
    }
    return <div>
        <PrimaryButton text="Apply Filter" onClick={openPanel} className={"filter-btn"} />
        <PrimaryButton text="Reset Filter" onClick={reset} className={"reset-btn"} />
        <Panel
            headerText="Filter"
            isOpen={isOpen}
            onDismiss={dismissPanel}
            isLightDismiss={true}
            closeButtonAriaLabel="Close"
        >
            <Text style={{ width: 200, paddingBottom: 80, paddingTop: 30 }} variant={"medium"}>Organization and Related People works only if no team is selected </Text>
            <ChoiceGroup defaultSelectedKey="1" style={{ marginBottom: 20 }} options={options} onChange={onChange} label="Users from" required={true} />
            <Departments data={data} teams={teams} team={team} />
        </Panel>
    </div>
}



