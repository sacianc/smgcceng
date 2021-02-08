import * as React from 'react';
import './draftMessages.scss';
import { initializeIcons } from 'office-ui-fabric-react/lib/Icons';
import { connect } from 'react-redux';
import { selectMessage, getDraftMessagesList, getMessagesList } from '../../actions';
import { getBaseUrl } from '../../configVariables';
import * as microsoftTeams from "@microsoft/teams-js";
import { Loader, List, Flex, Text } from '@stardust-ui/react';
import Overflow from '../OverFlow/draftMessageOverflow';

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
    id: string;
    title: string;
    lastSavedDate: string;
    recipients: string;
    acknowledgements?: string;
    reactions?: string;
    responses?: string;
    isRecurrence?: string;
}

export interface IMessageProps {
    messages: IMessage[];
    selectedMessage: any;
    searchText: string;
    selectMessage?: any;
    getDraftMessagesList?: any;
    getMessagesList?: any;
}

export interface IMessageState {
    message: IMessage[];
    itemsAccount: number;
    loader: boolean;
    teamsTeamId?: string;
    teamsChannelId?: string;
}

class DraftMessages extends React.Component<IMessageProps, IMessageState> {
    private interval: any;
    private isOpenTaskModuleAllowed: boolean;

    constructor(props: IMessageProps) {
        super(props);
        initializeIcons();
        this.isOpenTaskModuleAllowed = true;
        this.state = {
            message: props.messages,
            itemsAccount: this.props.messages.length,
            loader: true,
            teamsTeamId: "",
            teamsChannelId: "",
        };
    }

    public componentDidMount() {
        microsoftTeams.initialize();
        microsoftTeams.getContext((context) => {
            this.setState({
                teamsTeamId: context.teamId,
                teamsChannelId: context.channelId,
            });
        });

        this.props.getDraftMessagesList();
        this.interval = setInterval(() => {
            this.props.getDraftMessagesList();
        }, 60000);
    }

    public componentWillReceiveProps(nextProps: any) {
        this.setState({
            message: nextProps.messages,
            loader: false
        })
    }

    public componentWillUnmount() {
        clearInterval(this.interval);
    }

    private formatTime = (date: Date) => {
        return new Intl.DateTimeFormat('en-US', { hour: '2-digit', minute: '2-digit', hour12: true }).format(new Date(date)).toString().toUpperCase()
    };

    private formatNotificationDate = (notificationDate: string) => {
        if (notificationDate) {
            notificationDate = (new Date(notificationDate)).toLocaleString(navigator.language, { year: 'numeric', month: 'numeric', day: 'numeric', hour: 'numeric', minute: 'numeric', hour12: true });
            notificationDate = notificationDate.replace(',', '\xa0\xa0');
        }
        return notificationDate;
    }

    private _onFormatDate = (date: Date | null | undefined): string => {
        const shortMonths = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
        if (date != null) {
            return date.getMonth() + '/' + ('0' + date.getDate()).slice(-2) + '/' + (date.getFullYear());
        }
        return "";
    };

    public render(): JSX.Element {
        let keyCount = 0;
        const processItem = (message: any) => {
            keyCount++;
            const out = {
                key: keyCount,
                content: (
                    <Flex vAlign="center" fill gap="gap.small">
                        <Flex.Item size="size.quarter" variables={{ 'size.quarter': '58%' }} shrink={0} grow={1}>
                            <Text className="semiBold">{message.title}</Text>
                        </Flex.Item>
                        <Flex.Item size="size.quarter" variables={{ 'size.quarter': '11%' }}>
                            <Text>{message.isRecurrence ? "Yes" : "No"}</Text>
                        </Flex.Item>
                        <Flex.Item size="size.quarter" variables={{ 'size.quarter': '25%' }} >
                            <Text
                                truncated
                                content={this.formatNotificationDate(message.lastSavedDate)}
                            />
                        </Flex.Item>
                        <Flex.Item shrink={0}>
                            <Overflow message={message} title="" />
                        </Flex.Item>
                    </Flex>
                ),
                styles: { margin: '0.2rem 0.2rem 0 0' },
                onClick: (): void => {
                    let url = getBaseUrl() + "/newmessage/" + message.id;
                    this.onOpenTaskModule(null, url, "Edit message");
                },
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
        const allDraftMessages = [...label, ...outList];

        if (this.state.loader) {
            return (
                <Loader />
            );
        } else if (message.length === 0) {
            return (<div className="results">Hello, looks like you don't have any message in Draft. Please click the 'New Message' button to create new message.</div>);
        }
        else {
            return (
                <List selectable items={allDraftMessages} className="list" />
            );
        }
    }

    private processLabels = () => {
        const out = [{
            key: "labels",
            content: (
                <Flex vAlign="center" fill gap="gap.small">
                    <Flex.Item size="size.quarter" variables={{ 'size.quarter': '58%' }}>
                        <Text
                            truncated
                            weight="bold"
                            content="Title"
                        >
                        </Text>
                    </Flex.Item>
                    <Flex.Item size="size.quarter" variables={{ 'size.quarter': '11%' }}>
                        <Text
                            truncated
                            weight="bold"
                            content="Recurrence"
                        >
                        </Text>
                    </Flex.Item>
                    <Flex.Item size="size.quarter" variables={{ 'size.quarter': '25%' }}>
                        <Text
                            truncated
                            weight="bold"
                            content="Last saved"
                        >
                        </Text>
                    </Flex.Item>
                    <Flex.Item shrink={0}>
                        <Text
                            truncated
                            content=""
                        >
                        </Text>
                    </Flex.Item>
                </Flex>
            ),
            styles: { margin: '0.2rem 0.2rem 0 0' },
        }];
        return out;
    }

    private onOpenTaskModule = (event: any, url: string, title: string) => {
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
                this.props.getDraftMessagesList().then(() => {
                    this.props.getMessagesList();
                    this.isOpenTaskModuleAllowed = true;
                });
            };

            microsoftTeams.tasks.startTask(taskInfo, submitHandler);
        }
    }
}

const mapStateToProps = (state: any) => {
    return { messages: state.draftMessagesList, selectedMessage: state.selectedMessage };
}

export default connect(mapStateToProps, { selectMessage, getDraftMessagesList, getMessagesList })(DraftMessages);