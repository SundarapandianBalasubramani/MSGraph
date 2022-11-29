import { DefaultButton, PrimaryButton, SearchBox, Stack, StackItem } from "@fluentui/react";
import { MouseEventHandler } from "react";

export interface ISearch {
    onClick(key: string): void,
    letter: string,
    searchValue: string,
    onSearch(text: string): void,
    onChange(event?: React.ChangeEvent<HTMLInputElement>, newValue?: string): void
}

export const Search = (props: ISearch) => {

    const Buttons = () => {
        const buttons = [];
        for (let k = 65; k < 91; k++) {
            const str = String.fromCharCode(k);
            if (props.letter === str)
                buttons.push(<PrimaryButton onClick={() => { props.onClick(str) }} className="alphabetbutton" key={str}>{str}</PrimaryButton>);
            else
                buttons.push(<DefaultButton onClick={() => { props.onClick(str) }} className="alphabetbutton" key={str}>{str}</DefaultButton>);
        }
        return buttons;
    }
    return <Stack tokens={{ childrenGap: 20 }}>
        <StackItem>
            <SearchBox
                placeholder="Search"
                onSearch={props.onSearch}
                onChange={props.onChange}
            />
        </StackItem>
        <StackItem>
            <Stack horizontal wrap tokens={{ childrenGap: 15 }}>
                {Buttons()}
            </Stack>
        </StackItem>
    </Stack>
}
