import * as React from 'react';
import Messages from '../Messages/messages';
import DraftMessages from '../DraftMessages/draftMessages';
import ScheduledMessages from '../ScheduledMessages/scheduledMessages'
import './tabContainer.scss';
import * as microsoftTeams from "@microsoft/teams-js";
import { getBaseUrl } from '../../configVariables';
import { Button, FlexItem, Flex, Menu, Input, MenuItem } from '@stardust-ui/react';
import { getDraftMessagesList } from '../../actions';
import { connect } from 'react-redux';

interface ITaskInfo {
    title?: string;
    height?: number;
    width?: number;
    url?: string;
    card?: string;
    fallbackUrl?: string;
    completionBotId?: string;
}

export interface IActivityMenu {
    key: string,
    content: string,
    menuId: string,
}

export interface ITaskInfoProps {
    getDraftMessagesList?: any;
}

export interface ITabContainerState {
    url: string;
    activeMenuIndex: number;
    searchText: string;
    draftsCount: number;
}

class TabContainer extends React.Component<ITaskInfoProps, ITabContainerState> {
    constructor(props: ITaskInfoProps) {
        super(props);
        this.state = {
            url: getBaseUrl() + "/newmessage",
            activeMenuIndex: 0,
            searchText: "",
            draftsCount: 0,
        }
        this.escFunction = this.escFunction.bind(this);
    }

    public componentDidMount() {
        microsoftTeams.initialize();
        //- Handle the Esc key
        document.addEventListener("keydown", this.escFunction, false);
    }

    public componentWillReceiveProps(nextProps: any) {
        console.log(nextProps.messages.length);
        this.setState({
            draftsCount: nextProps.messages.length,
        })
    }

    public componentWillUnmount() {
        document.removeEventListener("keydown", this.escFunction, false);
    }

    public escFunction(event: any) {
        if (event.keyCode === 27 || (event.key === "Escape")) {
            microsoftTeams.tasks.submitTask();
        }
    }
    public onMenuClick = (value: any) => {
        this.setState({
            activeMenuIndex: value,
        })
    }

    public render(): JSX.Element {
        let menuItems: {}[] = [];
        menuItems.push(<MenuItem key="Drafts" content="Drafts" onClick={() => this.onMenuClick(0)}> Drafts ({this.state.draftsCount})
                        </MenuItem>);
        menuItems.push(<MenuItem key="Scheduled" content="Scheduled" onClick={() => this.onMenuClick(1)} >Scheduled
                        </MenuItem>);
        menuItems.push(<MenuItem key="Sent" content="Sent" onClick={() => this.onMenuClick(2)} >Sent
                        </MenuItem>);

        let activeComponent: {}[] = [];
        if (this.state.activeMenuIndex === 0) {
            activeComponent.push(<DraftMessages searchText={this.state.searchText}></DraftMessages>);
        }
        else if (this.state.activeMenuIndex === 1) {
            activeComponent.push(<ScheduledMessages searchText={this.state.searchText}></ScheduledMessages>);
        }
        else if (this.state.activeMenuIndex === 2) {
            activeComponent.push(<Messages searchText={this.state.searchText}></Messages>);
        }

        return (
            <div className="tabContainer">
                <Flex>
                    <Menu defaultActiveIndex={this.state.activeMenuIndex} className="tab-menu" items={menuItems} pointing primary />
                    <FlexItem className="newPostBtn" push>
                        <Input aria-label="Search" className="inputSearch" icon="search" placeholder="Search..." onChange={this.searchMessage} />
                    </FlexItem>
                    <Button content="New message" onClick={this.onNewMessage} primary />
                </Flex>
                <div className="messageContainer">
                    {activeComponent}
                </div>
            </div>
        );
    }

    private searchMessage = (e: React.SyntheticEvent<HTMLElement, Event>) => {
        let searchQuery = (e.target as HTMLInputElement).value;
        this.setState({
            searchText: searchQuery,
        })
    }

    public onNewMessage = () => {
        let taskInfo: ITaskInfo = {
            url: this.state.url,
            title: "New message",
            height: 530,
            width: 1000,
            fallbackUrl: this.state.url,
        }

        let submitHandler = (err: any, result: any) => {
            this.props.getDraftMessagesList();
        };

        microsoftTeams.tasks.startTask(taskInfo, submitHandler);
    }
}

const mapStateToProps = (state: any) => {
    return { messages: state.draftMessagesList };
}

export default connect(mapStateToProps, { getDraftMessagesList })(TabContainer);