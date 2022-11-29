import { Spinner, SpinnerSize, Text } from "@fluentui/react"
export interface ILoader {
    show: boolean
}
export const Loader = (props: ILoader) => {
    return props.show ? <div className="content">
        <div className="overlay">
            <div className="overlay-content">
                <Spinner size={SpinnerSize.large} />
                <Text>Loading</Text>
            </div>
        </div>
    </div> : <></>
}