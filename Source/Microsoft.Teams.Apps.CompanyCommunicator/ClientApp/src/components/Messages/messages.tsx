import * as React from 'react';
import { TooltipHost } from 'office-ui-fabric-react';
import './messages.scss';
import { initializeIcons } from 'office-ui-fabric-react/lib/Icons';
import { Icon, Loader, List, Flex, Text } from '@stardust-ui/react';
import { connect } from 'react-redux';
import { selectMessage, getMessagesList, getDraftMessagesList } from '../../actions';
import * as microsoftTeams from "@microsoft/teams-js";
import { getBaseUrl } from '../../configVariables';
import Overflow from '../OverFlow/sentMessageOverflow';

export interface ITaskInfo {
    title?: string;
    height?: number;
    width?: number;
    url?: string;
    card?: string;
    fallbackUrl?: string;
    completionBotId?: string;
}

export interface IMessage {
    title: string;
    sentDate: string;
    recipients: string;
    isRecurrence?: boolean;
    acknowledgements?: string;
    reactions?: string;
    responses?: string;
}

export interface IMessageProps {
    messagesList: IMessage[];
    searchText: string;
    selectMessage?: any;
    getMessagesList?: any;
    getDraftMessagesList?: any;
}

export interface IMessageState {
    message: IMessage[];
    loader: boolean;
}

class Messages extends React.Component<IMessageProps, IMessageState> {
    private interval: any;
    private isOpenTaskModuleAllowed: boolean;
    constructor(props: IMessageProps) {
        super(props);
        initializeIcons();
        this.isOpenTaskModuleAllowed = true;
        this.state = {
            message: this.props.messagesList,
            loader: true,
        };
        this.escFunction = this.escFunction.bind(this);
    }

    public componentDidMount() {
        microsoftTeams.initialize();
        this.props.getMessagesList();
        document.addEventListener("keydown", this.escFunction, false);
        this.interval = setInterval(() => {
            this.props.getMessagesList();
        }, 60000);
    }

    public componentWillUnmount() {
        document.removeEventListener("keydown", this.escFunction, false);
        clearInterval(this.interval);
    }

    public componentWillReceiveProps(nextProps: any) {
        if (this.props !== nextProps) {
            this.setState({
                message: nextProps.messagesList,
                loader: false
            });
        }
    }

    public render(): JSX.Element {
        let keyCount = 0;
        const processItem = (message: any) => {
            keyCount++;
            const out = {
                key: keyCount,
                content: this.messageContent(message),
                onClick: (): void => {
                    let url = getBaseUrl() + "/viewstatus/" + message.id;
                    this.onOpenTaskModule(null, url, "View status");
                },
                styles: { margin: '0.2rem 0.2rem 0 0' },
            };
            return out;
        };

        const label = this.processLabels();
        let searchText = this.props.searchText;
        let message: IMessage[];
        if (!searchText) // If Search text cleared
        {
            message = this.state.message;
        }
        else {
            message = this.state.message.filter((x: IMessage) => x.title.toLowerCase().includes(searchText.toLowerCase()));
        }

        const outList = message.map(processItem);
        const allMessages = [...label, ...outList];

        if (this.state.loader) {
            return (
                <Loader />
            );
        } else if (message.length === 0) {
            return (<div className="results">Hello, looks like you don't have any message in Sent. Please click the 'New Message' button to create new message.</div>);
        }
        else {
            return (
                <List selectable items={allMessages} className="list" />
            );
        }
    }

    private processLabels = () => {
        const out = [{
            key: "labels",
            content: (
                <Flex vAlign="center" fill gap="gap.small">
                    <Flex.Item size="size.quarter" variables={{ 'size.quarter': '40%' }} grow={1} >
                        <Text
                            truncated
                            weight="bold"
                            content="Title"
                        >
                        </Text>
                    </Flex.Item>
                    <Flex.Item size="size.quarter" variables={{ 'size.quarter': '23%' }} shrink={false}>
                        <Text
                            truncated
                            content="Recipients"
                            weight="bold"
                        >
                        </Text>
                    </Flex.Item>
                    <Flex.Item size="size.quarter" variables={{ 'size.quarter': '12%' }} shrink={false}>
                        <Text
                            truncated
                            content="Recurrence"
                        >
                        </Text>
                    </Flex.Item>
                    <Flex.Item size="size.quarter" variables={{ 'size.quarter': '15%' }} >
                        <Text
                            truncated
                            content="Sent"
                        >
                        </Text>
                    </Flex.Item>
                    <Flex.Item shrink={0} >
                        <Overflow title="" />
                    </Flex.Item>
                </Flex>
            ),
            styles: { margin: '0.2rem 0.2rem 0 0' },
        }];
        return out;
    }

    private renderSendingNumbersText = (message: any) => {
        if (message.isCompleted) {
            return (<Text />);
        } else {
            let currentNum = message.succeeded + message.failed + message.throttled;
            return (
                <Text
                    truncated
                    content={`Sending... ${currentNum} of ${message.totalMessageCount}`}
                >
                </Text>
            );
        }
    }

    private messageContent = (message: any) => {
        return (
            <Flex className="listContainer" vAlign="center" fill gap="gap.small">
                <Flex.Item size="size.quarter" variables={{ 'size.quarter': '42%' }} grow={1}>
                    <Text
                        truncated
                        className="semiBold"
                        content={message.title}
                    >
                    </Text>
                </Flex.Item>
                <Flex.Item size="size.quarter" variables={{ 'size.quarter': '24%' }} shrink={false}>
                    <div>
                        <TooltipHost content="Success" calloutProps={{ gapSpace: 0 }}>
                            <Icon name="stardust-checkmark" xSpacing="after" className="succeeded" outline />
                            <span className="semiBold">{message.succeeded}</span>
                        </TooltipHost>
                        <TooltipHost content="Failure" calloutProps={{ gapSpace: 0 }}>
                            <Icon name="stardust-close" xSpacing="both" className="failed" outline />
                            <span className="semiBold">{message.failed}</span>
                        </TooltipHost>
                        <TooltipHost content="Throttled" calloutProps={{ gapSpace: 0 }}>
                            <Icon name="exclamation-circle" xSpacing="both" className="throttled" outline />
                            <span className="semiBold">{message.throttled}</span>
                        </TooltipHost>
                        <TooltipHost content="Acknowledged" calloutProps={{ gapSpace: 0 }}>
                            <Icon name="eye" xSpacing="both" className="acknowledged" outline />
                            <span className="semiBold">{message.acknowledged}</span>
                        </TooltipHost>
                    </div>
                </Flex.Item>
                <Flex.Item size="size.quarter" variables={{ 'size.quarter': '12%' }} >
                    <Text
                        truncated
                        className="semiBold"
                        content={message.isRecurrence ? "Yes" : "No"}
                    />
                </Flex.Item>
                <Flex.Item size="size.quarter" variables={{ 'size.quarter': '15%' }} >
                    <Text
                        truncated
                        className="semiBold"
                        content={message.sentDate}
                    />
                </Flex.Item>
                <Flex.Item shrink={0}>
                    <Overflow message={message} title="" />
                </Flex.Item>
            </Flex>
        );
    }

    private escFunction = (event: any) => {
        if (event.keyCode === 27 || (event.key === "Escape")) {
            microsoftTeams.tasks.submitTask();
        }
    }

    public onOpenTaskModule = (event: any, url: string, title: string) => {
        if (this.isOpenTaskModuleAllowed) {
            this.isOpenTaskModuleAllowed = false;
            let taskInfo: ITaskInfo = {
                url: url,
                title: title,
                height: 530,
                width: 1000,
                fallbackUrl: url,
            }

            let submitHandler = (err: any, result: any) => {
                this.isOpenTaskModuleAllowed = true;
            };

            microsoftTeams.tasks.startTask(taskInfo, submitHandler);
        }
    }
}

const mapStateToProps = (state: any) => {
    return { messagesList: state.messagesList };
}

export default connect(mapStateToProps, { selectMessage, getMessagesList, getDraftMessagesList })(Messages);