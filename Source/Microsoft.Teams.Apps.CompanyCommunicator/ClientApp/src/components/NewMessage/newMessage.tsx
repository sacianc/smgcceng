import * as React from 'react';
import './newMessage.scss';
import './teamTheme.scss';
import { Input, TextArea, Radiobutton, RadiobuttonGroup } from 'msteams-ui-components-react';
import * as AdaptiveCards from "adaptivecards";
import { Button, Loader, Text, Flex, FlexItem, Divider, Checkbox, Dropdown, Icon } from '@stardust-ui/react';
import * as microsoftTeams from "@microsoft/teams-js";
import { RouteComponentProps } from 'react-router-dom';
import { getDraftNotification, getTeams, createDraftNotification, updateDraftNotification, getADGroups, getADGroupList, sendDraftNotification } from '../../apis/messageListApi';
import {
    getInitAdaptiveCard, setCardTitle, setCardImageLink, setCardSummary,
    setCardAuthor, setCardBtn
} from '../AdaptiveCard/adaptiveCard';
import { initializeIcons } from 'office-ui-fabric-react/lib/Icons';
import { Icon as IconFabric } from 'office-ui-fabric-react/lib/Icon';
import { getBaseUrl } from '../../configVariables';
import { isNullOrEmpty } from 'adaptivecards';
import { isNullOrUndefined } from 'util';
import { mergeStyleSets } from 'office-ui-fabric-react/lib/Styling';
import { DatePicker, DayOfWeek, IDatePickerStrings } from 'office-ui-fabric-react/lib/DatePicker';
import { Dropdown as FabricDropdown, IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown';

const DayPickerStrings: IDatePickerStrings = {
    months: ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'],
    shortMonths: ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'],
    days: ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'],
    shortDays: ['S', 'M', 'T', 'W', 'T', 'F', 'S'],

    goToToday: '',
    prevMonthAriaLabel: 'Go to previous month',
    nextMonthAriaLabel: 'Go to next month',
    prevYearAriaLabel: 'Go to previous year',
    nextYearAriaLabel: 'Go to next year',
    closeButtonAriaLabel: 'Close date picker'
};
const repeatsValues = ['Every weekday (Mon-Fri)', 'Daily', 'Weekly', 'Monthly', 'Yearly', 'Custom'];
const repeatsdropdownValues: dropdownItem[] = [];

repeatsValues.forEach((element) => {
    repeatsdropdownValues.push({
        header: element.valueOf(),
        team: {
            id: element.valueOf()
        }
    });
});

const repeatFrequency = ['Day', 'Week', 'Month'];
const repeatFrequencydropdownValues: dropdownItem[] = [];
repeatFrequency.forEach((element) => {
    repeatFrequencydropdownValues.push({
        header: element.valueOf(),
        team: {
            id: element.valueOf()
        }
    });
});

const controlClass = mergeStyleSets({
    control: {
        margin: '0 0 15px 0',
        maxWidth: '300px'
    }
});

type dropdownItem = {
    header: string,
    team: {
        id: string,
    },
}

export interface ADGroup {
    id: string,
    displayName: string,
}

export interface IDraftMessage {
    id?: string,
    title: string,
    imageLink?: string,
    summary?: string,
    author: string,
    buttonTitle?: string,
    buttonLink?: string,
    teams: any[],
    rosters: any[],
    adGroups: any[],
    allUsers: boolean,
    buttonTitle2?: string,
    buttonLink2?: string,
    IsScheduled: boolean,
    scheduleDate?: Date,
    isRecurrence: boolean,
    repeats?: string,
    repeatFor?: number,
    repeatFrequency?: string,
    weekSelection?: string,
    repeatStartDate?: Date,
    repeatEndDate?: Date,
}

export interface formState {
    title: string,
    summary?: string,
    btnLink?: string,
    imageLink?: string,
    btnTitle?: string,
    author: string,
    card?: any,
    page: string,
    teamsOptionSelected: boolean,
    rostersOptionSelected: boolean,
    adGroupOptionSelected: boolean,
    allUsersOptionSelected: boolean,
    teams?: any[],
    adGroups?: any[],
    exists?: boolean,
    messageId: string,
    loader: boolean,
    selectedTeamsNum: number,
    selectedRostersNum: number,
    selectedADGroupsNum: number,
    selectedRadioBtn: string,
    selectedTeams: dropdownItem[],
    selectedRosters: dropdownItem[],
    selectedADGroups: dropdownItem[],
    resultADGroups: dropdownItem[],
    errorImageUrlMessage: string,
    errorButtonUrlMessage: string,
    btnLink2?: string,
    btnTitle2?: string,
    errorButtonUrlMessage2: string,
    IsScheduled: boolean,
    scheduleDate: Date,
    defaultSelectedIndexStartTime: string,
    isRecurrenceEnabled: boolean,
    repeats: string,
    repeatsStartDate: Date,
    repeatsFor: number,
    repeatFrequency: string,
    weekSelection: string,
    repeatsEndDate: Date,
    mondaySelection: boolean,
    tuesdaySelection: boolean,
    webnesdaySelection: boolean,
    thursdaySelection: boolean,
    fridaySelection: boolean,
    saturdaySelection: boolean,
    sundaySelection: boolean,
    isRepeatTimeValid: boolean,
    isWeeksSelectionValid: boolean,
}

export interface INewMessageProps extends RouteComponentProps {
    getDraftMessagesList?: any;
}

const todayDate: Date = new Date();

export default class NewMessage extends React.Component<INewMessageProps, formState> {
    private card: any;

    constructor(props: INewMessageProps) {
        super(props);
        initializeIcons();
        let currentDate = this.datetime(new Date(), "00:00 AM");
        let scheduleDate = this.datetime(new Date(), "08:30 AM");
        this.card = getInitAdaptiveCard();
        this.setDefaultCard(this.card);

        this.state = {
            title: "",
            summary: "",
            author: "",
            btnLink: "",
            imageLink: "",
            btnTitle: "",
            card: this.card,
            page: "CardCreation",
            teamsOptionSelected: false,
            rostersOptionSelected: false,
            adGroupOptionSelected: true,
            allUsersOptionSelected: false,
            messageId: "",
            loader: true,
            selectedTeamsNum: 0,
            selectedRostersNum: 0,
            selectedADGroupsNum: 0,
            selectedRadioBtn: "adgroups",
            selectedTeams: [],
            selectedRosters: [],
            selectedADGroups: [],
            resultADGroups: [],
            errorImageUrlMessage: "",
            errorButtonUrlMessage: "",
            btnLink2: "",
            btnTitle2: "",
            errorButtonUrlMessage2: "",
            IsScheduled: false,
            scheduleDate: scheduleDate,
            defaultSelectedIndexStartTime: "08:30 AM",
            isRecurrenceEnabled: false,
            repeats: repeatsValues[0],
            repeatsStartDate: currentDate,
            repeatsFor: 0,
            repeatFrequency: repeatFrequency[0],
            weekSelection: "",
            repeatsEndDate: currentDate,
            mondaySelection: false,
            tuesdaySelection: false,
            webnesdaySelection: false,
            thursdaySelection: false,
            fridaySelection: false,
            saturdaySelection: false,
            sundaySelection: false,
            isRepeatTimeValid: true,
            isWeeksSelectionValid: true,
        }
    }

    public async componentDidMount() {
        microsoftTeams.initialize();
        //- Handle the Esc key
        document.addEventListener("keydown", this.escFunction, false);
        let params = this.props.match.params;
        this.getTeamList().then(() => {
            if ('id' in params) {
                let id = params['id'];
                this.getItem(id).then(() => {
                    const selectedTeams = this.makeDropdownItemList(this.state.selectedTeams, this.state.teams);
                    const selectedRosters = this.makeDropdownItemList(this.state.selectedRosters, this.state.teams);
                    const selectedADGroups = this.makeDropdownItemList(this.state.selectedADGroups, this.state.adGroups);
                    this.setState({
                        exists: true,
                        messageId: id,
                        selectedTeams: selectedTeams,
                        selectedRosters: selectedRosters,
                        selectedADGroups: selectedADGroups,
                    })
                });
            } else {
                this.setState({
                    exists: false,
                    loader: false
                }, () => {
                    let adaptiveCard = new AdaptiveCards.AdaptiveCard();
                    adaptiveCard.parse(this.state.card);
                    let renderedCard = adaptiveCard.render();
                    document.getElementsByClassName('adaptiveCardContainer')[0].appendChild(renderedCard);
                    if (this.state.btnLink) {
                        let link = this.state.btnLink;
                        adaptiveCard.onExecuteAction = function (action) { window.open(link, '_blank'); };
                    }
                })
            }
        });
    }

    private makeDropdownItemList = (items: any[], fromItems: any[] | undefined) => {
        const dropdownItemList: dropdownItem[] = [];
        items.forEach(element =>
            dropdownItemList.push(
                typeof element !== "string" ? element : {
                    header: fromItems!.find(x => x.teamId === element).name,
                    team: {
                        id: element
                    }
                })
        );
        return dropdownItemList;
    }


    public setDefaultCard = (card: any) => {
        setCardTitle(card, "Title");
        let imgUrl = getBaseUrl() + "/image/imagePlaceholder.png";
        setCardImageLink(card, imgUrl);
        setCardSummary(card, "Summary");
        setCardAuthor(card, "- Author");
        setCardBtn(card, "Primary Button title", "https://adaptivecards.io", "Secondary Button title", "https://forms.office.com");
    }

    private getTeamList = async () => {
        try {
            const response = await getTeams();
            this.setState({
                teams: response.data
            });
        } catch (error) {
            return error;
        }
    }

    private getADGroupList = async (adGroupIds: ADGroup[]) => {
        console.log(adGroupIds);
        try {
            const response = await getADGroupList(adGroupIds);

            if (response.status === 200 && response.data != null) {
                let adGroupsList: ADGroup[] = response.data;
                this.setState({
                    adGroups: adGroupsList
                });
            }

        } catch (error) {
            return error;
        }
    }

    private getItem = async (id: number) => {
        try {
            const response = await getDraftNotification(id);
            const draftMessageDetail = response.data;
            const selectedRadioButton = draftMessageDetail.rosters.length > 0 ? "rosters" : draftMessageDetail.allUsers ? "allUsers" : draftMessageDetail.adGroups.length > 0 ? "adgroups" : "teams";

            if (selectedRadioButton === "adgroups" && draftMessageDetail.adGroups != null) {
                let adGroupsList: ADGroup[] = [];
                for (var i = 0; i < draftMessageDetail.adGroups.length; i++) {
                    let adGroup: ADGroup = {
                        id: draftMessageDetail.adGroups[i],
                        displayName: ""
                    };

                    adGroupsList.push(adGroup);

                }
                await this.getADGroupList(adGroupsList);
            }

            let scheduleDateValue = new Date(draftMessageDetail.scheduleDate);
            let repeatsStartDateValue = new Date(draftMessageDetail.repeatStartDate);
            let repeatEndDateValue = new Date(draftMessageDetail.repeatEndDate);
            let currentDate = this.datetime(new Date(), "00:00 AM");

            if (!draftMessageDetail.isScheduled) {
                scheduleDateValue = currentDate;
            }
            if (!draftMessageDetail.isRecurrence) {
                repeatsStartDateValue = currentDate;
                repeatEndDateValue = currentDate;
            }

            this.setState({
                teamsOptionSelected: draftMessageDetail.teams.length > 0,
                selectedTeamsNum: draftMessageDetail.teams.length,
                rostersOptionSelected: draftMessageDetail.rosters.length > 0,
                adGroupOptionSelected: draftMessageDetail.adGroups.length > 0,
                selectedRostersNum: draftMessageDetail.rosters.length,
                selectedADGroupsNum: draftMessageDetail.adGroups.length,
                selectedRadioBtn: selectedRadioButton,
                selectedTeams: draftMessageDetail.teams,
                selectedRosters: draftMessageDetail.rosters,
                selectedADGroups: draftMessageDetail.adGroups,
                messageId: draftMessageDetail.id,
                exists: true,
                IsScheduled: draftMessageDetail.isScheduled,
                scheduleDate: scheduleDateValue,
                defaultSelectedIndexStartTime: this.formatTime(draftMessageDetail.scheduleDate),
                isRecurrenceEnabled: draftMessageDetail.isRecurrence,
                repeats: draftMessageDetail.repeats,
                repeatsFor: draftMessageDetail.repeatFor,
                repeatFrequency: draftMessageDetail.repeatFrequency,
                weekSelection: draftMessageDetail.weekSelection,
                repeatsStartDate: repeatsStartDateValue,
                repeatsEndDate: repeatEndDateValue,
            });

            setCardTitle(this.card, draftMessageDetail.title);
            setCardImageLink(this.card, draftMessageDetail.imageLink);
            setCardSummary(this.card, draftMessageDetail.summary);
            setCardAuthor(this.card, draftMessageDetail.author);
            setCardBtn(this.card, draftMessageDetail.buttonTitle, draftMessageDetail.buttonLink, draftMessageDetail.buttonTitle2, draftMessageDetail.buttonLink2);

            this.setState({
                title: draftMessageDetail.title,
                summary: draftMessageDetail.summary,
                btnLink: draftMessageDetail.buttonLink,
                imageLink: draftMessageDetail.imageLink,
                btnTitle: draftMessageDetail.buttonTitle,
                author: draftMessageDetail.author,
                allUsersOptionSelected: draftMessageDetail.allUsers,
                btnLink2: draftMessageDetail.buttonLink2,
                btnTitle2: draftMessageDetail.buttonTitle2,
                loader: false
            }, () => {
                this.updateCard();
            });
        } catch (error) {
            return error;
        }
    }

    public componentWillUnmount() {
        document.removeEventListener("keydown", this.escFunction, false);
    }

    private formatNotificationDate = (notificationDate: string) => {
        if (notificationDate) {
            notificationDate = (new Date(notificationDate)).toLocaleString(navigator.language, { year: 'numeric', month: 'numeric', day: 'numeric', hour: 'numeric', minute: 'numeric', hour12: true });
            notificationDate = notificationDate.replace(',', '\xa0\xa0');
        }
        return notificationDate;
    }

    public render(): JSX.Element {
        let minInterval = 30; //minutes interval
        const times = []; // time array
        let startTime = 0; // start time
        let ap = ['AM', 'PM']; // AM-PM

        //loop to increment the time and push results in array
        for (let i = 0; startTime < 24 * 60; i++) {
            let hh = Math.floor(startTime / 60); // getting hours of day in 0-24 format
            let mm = (startTime % 60); // getting minutes of the hour in 0-55 format
            times[i] = ("0" + (hh % 12)).slice(-2) + ':' + ("0" + mm).slice(-2) + ' ' + ap[Math.floor(hh / 12)]; // pushing data in array in [00:00 - 12:00 AM/PM format]
            startTime = startTime + minInterval;
        }
        const timeData: IDropdownOption[] = [];

        for (let i = 0; i < times.length; i++) {
            timeData.push({ key: times[i], text: times[i] });
        }

        let scheduleCheckbox: {} = {};
        let recurrenceCheckbox: {} = {};

        if (this.state.IsScheduled) {
            scheduleCheckbox = <Checkbox toggle defaultChecked onClick={this.onScheduleToggle} label="On" />;
        }
        else {
            scheduleCheckbox = <Checkbox toggle onClick={this.onScheduleToggle} label="Off" />;
        }
        if (this.state.isRecurrenceEnabled) {
            recurrenceCheckbox = <Checkbox toggle defaultChecked onClick={this.onRecurrenceToggle} label="On" />;
        }
        else {
            recurrenceCheckbox = <Checkbox toggle={!this.state.isRecurrenceEnabled} onClick={this.onRecurrenceToggle} label="Off" />;
        }

        let recurrenceMessage: string = "";
        if (this.state.isRecurrenceEnabled) {
            recurrenceMessage = "Occurs ";
            let repeats: string = this.state.repeats ? this.state.repeats : "";
            let repeatFrequency: string = this.state.repeatFrequency ? this.state.repeatFrequency : "";

            if (this.state.repeats !== "Custom") {
                recurrenceMessage += repeats.toLowerCase();
            }
            else if (this.state.repeats === "Custom") {
                if (this.state.repeatFrequency === "Day" || this.state.repeatFrequency === "Month") {
                    recurrenceMessage += "every " + this.state.repeatsFor + " " + repeatFrequency.toLowerCase();
                }
                else if (this.state.repeatFrequency === "Week") {
                    let weeks = "";
                    let weekSelection: string = this.state.weekSelection ? this.state.weekSelection : "";
                    if (weekSelection.indexOf('0') !== -1) {
                        weeks = "Monday,";
                    }
                    if (weekSelection.indexOf('1') !== -1) {
                        weeks += "Tuesday,";
                    }
                    if (weekSelection.indexOf('2') !== -1) {
                        weeks += "Wednesday,";
                    }
                    if (weekSelection.indexOf('3') !== -1) {
                        weeks += "Thursday,";
                    }
                    if (weekSelection.indexOf('4') !== -1) {
                        weeks += "Friday,";
                    }
                    if (weekSelection.indexOf('5') !== -1) {
                        weeks += "Saturday,";
                    }
                    if (weekSelection.indexOf('6') !== -1) {
                        weeks += "Sunday,";
                    }

                    recurrenceMessage += "every " + this.state.repeatsFor + " " + repeatFrequency.toLowerCase() + "(" + weeks.substring(0, weeks.length - 1) + ")";
                }
            }
            let repeatStartDate: string = this._onFormatDate2(this.state.repeatsStartDate);
            recurrenceMessage += " starting " + repeatStartDate;//.slice(repeatStartDate.length - 8)
        }

        if (this.state.loader) {
            return (
                <div className="Loader">
                    <Loader />
                </div>
            );
        } else {
            if (this.state.page === "CardCreation") {
                return (
                    <div className="taskModule">
                        <div className="formContainer">
                            <div className="formContentContainer" >
                                <Input
                                    className="inputField"
                                    value={this.state.title}
                                    label="Title"
                                    placeholder="Title (required)"
                                    onChange={this.onTitleChanged}
                                    autoComplete="off"
                                    required
                                />

                                <Input
                                    className="inputField"
                                    value={this.state.imageLink}
                                    label="Image URL"
                                    placeholder="Image URL"
                                    onChange={this.onImageLinkChanged}
                                    errorLabel={this.state.errorImageUrlMessage}
                                    autoComplete="off"
                                />

                                <TextArea
                                    className="inputField textArea"
                                    autoFocus
                                    placeholder="Summary"
                                    label="Summary"
                                    value={this.state.summary}
                                    onChange={this.onSummaryChanged}
                                />

                                <Input
                                    className="inputField"
                                    value={this.state.author}
                                    label="Author"
                                    placeholder="Author"
                                    onChange={this.onAuthorChanged}
                                    autoComplete="off"
                                />
                                <Flex>
                                    <FlexItem >
                                        <Input
                                            className="inputField Urltitle margin-right-0"
                                            value={this.state.btnTitle}
                                            label="Primary button title"
                                            placeholder="Primary button title"
                                            onChange={this.onBtnTitleChanged}
                                            autoComplete="off"
                                        />
                                    </FlexItem>
                                    <FlexItem >
                                        <Input
                                            className="inputField UrlInputLg margin-left-20 margin-right-0"
                                            value={this.state.btnLink}
                                            label="Primary button URL"
                                            placeholder="Primary button URL"
                                            onChange={this.onBtnLinkChanged}
                                            errorLabel={this.state.errorButtonUrlMessage}
                                            autoComplete="off"
                                        />
                                    </FlexItem>
                                </Flex>
                                <Flex>
                                    <FlexItem>
                                        <Input
                                            className="inputField Urltitle margin-right-0"
                                            value={this.state.btnTitle2}
                                            label="Secondary button title"
                                            placeholder="Secondary button title"
                                            onChange={this.onBtnTitleChanged2}
                                            autoComplete="off"
                                        />
                                    </FlexItem>
                                    <FlexItem>
                                        <Input
                                            className="inputField UrlInputLg margin-left-20  margin-right-0"
                                            value={this.state.btnLink2}
                                            label="Secondary button URL"
                                            placeholder="Secondary button URL"
                                            onChange={this.onBtnLinkChanged2}
                                            errorLabel={this.state.errorButtonUrlMessage2}
                                            autoComplete="off"
                                        />
                                    </FlexItem>
                                </Flex>
                            </div>
                            <div className="adaptiveCardContainer">
                            </div>
                        </div>

                        <div className="footerContainer">
                            <div className="buttonContainer">
                                <Button content="Next" disabled={this.isNextBtnDisabled()} id="saveBtn" onClick={this.onNext} primary />
                            </div>
                        </div>
                    </div>
                );
            }
            else if (this.state.page === "AudienceSelection") {
                return (
                    <div className="taskModule">
                        <div className="formContainer">
                            <div className="formContentContainer" >
                                <h3>Select how you want to send this message</h3>
                                <RadiobuttonGroup
                                    className="radioBtns"
                                    value={this.state.selectedRadioBtn}
                                    onSelected={this.onGroupSelected}
                                >
                                    <Radiobutton className="mt-2 mb-2" name="grouped" value="teams" label="Send to General channel of specific teams" />
                                    <Dropdown
                                        hidden={!this.state.teamsOptionSelected}
                                        placeholder="Select team(s)"
                                        search
                                        multiple
                                        items={this.getItems()}
                                        value={this.state.selectedTeams}
                                        onSelectedChange={this.onTeamsChange}
                                        noResultsMessage="We couldn't find any matches."
                                        className="dropDownmrg"
                                    />
                                    <Radiobutton className="mt-2 mb-2" name="grouped" value="rosters" label="Send in chat to specific person" />
                                    <Dropdown
                                        hidden={!this.state.rostersOptionSelected}
                                        placeholder="Choose team(s) members"
                                        search
                                        multiple
                                        items={this.getItems()}
                                        value={this.state.selectedRosters}
                                        onSelectedChange={this.onRostersChange}
                                        noResultsMessage="We couldn't find any matches."
                                        className="dropDownmrg"
                                    />
                                    <Radiobutton className="mt-2 mb-2" name="grouped" value="allUsers" label="Send in chat to everyone" />
                                    <div className={this.state.selectedRadioBtn === "allUsers" ? "" : "hide"}>
                                        <div className="notemsg">
                                            <Text error content="Note: This option sends the message to everyone in your org who has access to the app." />
                                        </div>
                                    </div>
                                    <Radiobutton className="mt-2 mb-2" name="grouped" value="adgroups" label="Send to distribution group" />
                                    <Dropdown
                                        hidden={!this.state.adGroupOptionSelected}
                                        placeholder="Please enter distribution group alias"
                                        search
                                        multiple
                                        items={this.state.resultADGroups}
                                        value={this.state.selectedADGroups}
                                        onSelectedChange={this.onADGroupChange}
                                        onSearchQueryChange={this.getADGroupItems}
                                        noResultsMessage="We couldn't find any matches."
                                        className="dropDownmrg"
                                    />
                                </RadiobuttonGroup>
                                <div className="contentwrapper">
                                    <Divider className="pt-2 pb-2" />
                                    <Flex gap="gap.large" className="mb-2">
                                        <Text content="Schedule send" weight="bold" />
                                        <div className="ml-auto"> {scheduleCheckbox}</div>
                                    </Flex>
                                    <Text content={this.state.IsScheduled ? "Choose a date and time" : "To send it later turn on the toggle. You will be asked to choose a date and time"} />
                                    <Flex gap="gap.small" >
                                        <DatePicker
                                            hidden={!this.state.IsScheduled}
                                            className={controlClass.control}
                                            strings={DayPickerStrings}
                                            showWeekNumbers={false}
                                            firstWeekOfYear={1}
                                            showMonthPickerAsOverlay={true}
                                            placeholder="Select a date..."
                                            ariaLabel="Select a date"
                                            formatDate={this._onFormatDate}
                                            minDate={todayDate}
                                            onSelectDate={this._onSelectScheduleDate}
                                            value={this.state.scheduleDate}
                                        />
                                        <FabricDropdown
                                            hidden={!this.state.IsScheduled}
                                            contentEditable
                                            onInput={this.onStartTimeChange}
                                            defaultSelectedKey={this.state.defaultSelectedIndexStartTime}
                                            options={timeData}
                                            placeholder="select start time"
                                            onChange={this.setStartTime}
                                        />
                                    </Flex>
                                    <Divider className="pt-2 pb-2" />
                                    <Flex gap="gap.large" className="mb-2">
                                        <Text content="Recurrence message" weight="bold" />
                                        <div className="ml-auto">{recurrenceCheckbox}</div>
                                    </Flex>
                                    <Text content={this.state.isRecurrenceEnabled ? "Choose a option" : "To set this message as a recurrence turn on the toggle"} />
                                    <Dropdown
                                        className="mb-3"
                                        hidden={!this.state.isRecurrenceEnabled}
                                        placeholder="Repeats"
                                        items={repeatsdropdownValues}
                                        value={this.state.repeats ? this.state.repeats : repeatsdropdownValues[0]}
                                        onSelectedChange={this.onRepeatChange}
                                        fluid
                                    />
                                    <Flex gap="gap.small">
                                        <Text
                                            content="Start"
                                            hidden={!this.state.isRecurrenceEnabled}
                                            className="pt-2"
                                        />
                                        <DatePicker
                                            hidden={!this.state.isRecurrenceEnabled}
                                            className={controlClass.control}
                                            strings={DayPickerStrings}
                                            showWeekNumbers={false}
                                            firstWeekOfYear={1}
                                            showMonthPickerAsOverlay={true}
                                            placeholder="Select a date..."
                                            ariaLabel="Select a date"
                                            formatDate={this._onFormatDate}
                                            minDate={todayDate}
                                            onSelectDate={this._onSelectStartDate}
                                            value={this.state.repeatsStartDate}
                                        />
                                        <Text
                                            hidden={!this.state.isRecurrenceEnabled}
                                            content="End"
                                            className="pt-2"
                                        />
                                        <DatePicker
                                            hidden={!this.state.isRecurrenceEnabled}
                                            className={controlClass.control}
                                            strings={DayPickerStrings}
                                            showWeekNumbers={false}
                                            firstWeekOfYear={1}
                                            showMonthPickerAsOverlay={true}
                                            placeholder="Select date..."
                                            ariaLabel="End date"
                                            formatDate={this._onFormatDate}
                                            minDate={todayDate}
                                            onSelectDate={this._onSelectEndDate}
                                            value={this.state.repeatsEndDate}
                                        />
                                    </Flex>
                                    <Flex gap="gap.small" className="customRecurrence">
                                        <Text
                                            content="Repeats every"
                                            className="customRecurrencetxt"
                                            hidden={!this.state.isRecurrenceEnabled || this.state.repeats !== repeatsValues[5]}
                                        />
                                        <Input
                                            hidden={!this.state.isRecurrenceEnabled || this.state.repeats.toString() !== repeatsValues[5]}
                                            className="inputField input-xs"
                                            value={this.state.repeatsFor}
                                            onChange={this.onRepeatForChanged}
                                            autoComplete="off"
                                        />
                                        <Dropdown
                                            hidden={!this.state.isRecurrenceEnabled || this.state.repeats.toString() !== repeatsValues[5]}
                                            items={repeatFrequencydropdownValues}
                                            value={this.state.repeatFrequency}
                                            onSelectedChange={this.onRepeatFrequencyChange}
                                            className="customRecurrence-drop"
                                        />
                                    </Flex>
                                    <Flex className="weekdaywrapper">
                                        <span hidden={!this.state.isRecurrenceEnabled || this.state.repeats.toString() !== repeatsValues[5] || this.state.repeatFrequency !== repeatFrequency[1]}>
                                            <FlexItem push>
                                                <Checkbox label="M" onClick={(e) => this.onWeekChange(e, 0)} />
                                            </FlexItem>
                                            <Checkbox label="T" onClick={(e) => this.onWeekChange(e, 1)} />
                                            <Checkbox label="W" onClick={(e) => this.onWeekChange(e, 2)} />
                                            <Checkbox label="T" onClick={(e) => this.onWeekChange(e, 3)} />
                                            <Checkbox label="F" onClick={(e) => this.onWeekChange(e, 4)} />
                                            <Checkbox label="S" onClick={(e) => this.onWeekChange(e, 5)} />
                                            <Checkbox label="S" onClick={(e) => this.onWeekChange(e, 6)} />
                                        </span>
                                    </Flex>
                                    <Flex gap="gap.small">
                                        <Text
                                            color="red"
                                            hidden={this.state.isRepeatTimeValid && this.state.isWeeksSelectionValid}
                                            content={!this.state.isRepeatTimeValid ? 'End date should be greater than start date' : (!this.state.isWeeksSelectionValid ? "Atleast one week must be selected" : "")} />
                                    </Flex>
                                    <Flex gap="gap.small">
                                        <FlexItem>
                                            <IconFabric iconName='Sync' hidden={!this.state.isRecurrenceEnabled} />
                                        </FlexItem>
                                        <FlexItem>
                                            <Text
                                                hidden={!this.state.isRecurrenceEnabled}
                                                content={recurrenceMessage} />
                                        </FlexItem>
                                    </Flex>
                                </div>
                            </div>
                            <div className="adaptiveCardContainer">
                            </div>
                        </div>

                        <div className="footerContainer">
                            <div className="buttonwrap">
                                <Flex gap="gap.small">
                                    <Button className="plainbtn" text content="< Back" onClick={this.onBack} />
                                    <FlexItem push>
                                        <Button content="Save as draft" disabled={this.isSaveBtnDisabled()} id="saveBtn" onClick={this.onSave} />
                                    </FlexItem>
                                    <Loader id="sendingLoader" className="hiddenLoader sendingLoader" size="smallest" label="Preparing message" labelPosition="end" />
                                    <Button content="Send" disabled={this.isSaveBtnDisabled()} id="saveBtn" onClick={this.onSend} primary />
                                </Flex>
                            </div>
                        </div>
                    </div>
                );
            } else {
                return (<div>Error</div>);
            }
        }
    }

    private formatTime = (date: Date) => {
        return new Intl.DateTimeFormat('en-US', { hour: '2-digit', minute: '2-digit', hour12: true }).format(new Date(date)).toString().toUpperCase()
    };

    private onScheduleToggle = () => {
        this.setState({
            IsScheduled: !this.state.IsScheduled,
        })
    }

    private onRecurrenceToggle = () => {
        this.setState({
            isRecurrenceEnabled: !this.state.isRecurrenceEnabled,
        })
    }

    private onWeekChange = (e: React.SyntheticEvent<HTMLElement, Event>, weekday: number) => {
        let mondaySelection = this.state.mondaySelection;
        let tuesdaySelection = this.state.tuesdaySelection;
        let webnesdaySelection = this.state.webnesdaySelection;
        let thursdaySelection = this.state.thursdaySelection;
        let fridaySelection = this.state.fridaySelection;
        let saturdaySelection = this.state.saturdaySelection;
        let sundaySelection = this.state.sundaySelection;

        if (weekday === 0) {
            mondaySelection = !this.state.mondaySelection;
        }
        else if (weekday === 1) {
            tuesdaySelection = !this.state.tuesdaySelection;
        }
        else if (weekday === 2) {
            webnesdaySelection = !this.state.webnesdaySelection;
        }
        else if (weekday === 3) {
            thursdaySelection = !this.state.thursdaySelection;
        }
        else if (weekday === 4) {
            fridaySelection = !this.state.fridaySelection;
        }
        else if (weekday === 5) {
            saturdaySelection = !this.state.saturdaySelection;
        }
        else if (weekday === 6) {
            sundaySelection = !this.state.sundaySelection;
        }

        this.setState({
            mondaySelection: mondaySelection,
            tuesdaySelection: tuesdaySelection,
            webnesdaySelection: webnesdaySelection,
            thursdaySelection: thursdaySelection,
            fridaySelection: fridaySelection,
            saturdaySelection: saturdaySelection,
            sundaySelection: sundaySelection
        },
            () => {
                if (this.state.repeats === repeatsValues[5] && this.state.repeatFrequency === "Week" && !this.state.mondaySelection && !this.state.tuesdaySelection && !this.state.tuesdaySelection && !this.state.webnesdaySelection && !this.state.thursdaySelection && !this.state.fridaySelection && !this.state.saturdaySelection && !this.state.sundaySelection) {
                    this.setState({
                        isWeeksSelectionValid: false,
                    })
                }
                else {
                    this.setState({
                        isWeeksSelectionValid: true,
                    })
                }
            });

    }

    private _onFormatDate = (date: Date | null | undefined): string => {
        const shortMonths = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
        if (date != null) {
            return shortMonths[date.getMonth()] + ' ' + ('0' + date.getDate()).slice(-2) + ', ' + (date.getFullYear());
        }
        return "";
    };

    private _onFormatDate2 = (date: Date | null | undefined): string => {
        if (date != null) {
            return date.getMonth() + 1 + '/' + date.getDate() + '/' + (date.getFullYear());
        }
        return "";
    };

    private _onSelectStartDate = (date: Date | null | undefined): void => {
        if (date != null) {
            this.setState(
                {
                    repeatsStartDate: date
                },
                () => {
                    if (this.state.repeatsStartDate > this.state.repeatsEndDate) {
                        this.setState({
                            isRepeatTimeValid: false,
                        })
                    }
                    else {
                        this.setState({
                            isRepeatTimeValid: true,
                        })
                    }
                });
        }
    };

    private _onSelectEndDate = (date: Date | null | undefined): void => {
        if (date != null) {
            this.setState(
                {
                    repeatsEndDate: date
                },
                () => {
                    if (this.state.repeatsStartDate > this.state.repeatsEndDate) {
                        this.setState({
                            isRepeatTimeValid: false,
                        })
                    }
                    else {
                        this.setState({
                            isRepeatTimeValid: true,
                        })
                    }
                });
        }
    };

    private _onSelectScheduleDate = (date: Date | null | undefined): void => {

        if (date != null) {
            let dateValue = this.datetime(date, this.state.defaultSelectedIndexStartTime);
            this.setState({ scheduleDate: dateValue });
        }
    };

    private onStartTimeChange = (e: React.SyntheticEvent<HTMLElement, Event>) => {
        let startTime = (e.target as HTMLInputElement).innerText;
        let regExp = /^(1[0-2]|0?[1-9]):[0-5][0-9] (AM|PM)$/i;
        let extractedTime = startTime.toLowerCase().split("m")[0].toLowerCase() + "m";
        if (regExp.test(extractedTime)) {
            let date = this.datetime(this.state.scheduleDate, startTime);
            this.setState({
                scheduleDate: date,
                defaultSelectedIndexStartTime: startTime,
            })

        }
    }

    private onRepeatFrequenceChange = (event: any, itemsData: any) => {
        this.setState({
            repeatFrequency: itemsData.value,
        })
    }

    private setStartTime = (e: React.SyntheticEvent<HTMLElement>, option?: IDropdownOption, index?: number) => {
        console.log(option!.text);
        if (option != null) {
            let date = this.datetime(this.state.scheduleDate, option!.text);
            console.log(date);
            this.setState({
                defaultSelectedIndexStartTime: option!.text,
                scheduleDate: date,
            });
        }
    }

    private datetime = (date: Date, time: string) => {
        let timeFormat = this.convertTime12to24(time).split(':');
        return new Date(date.getFullYear(), date.getMonth(), date.getDate(), parseInt(timeFormat[0]), parseInt(timeFormat[1]))
    }

    private convertTime12to24 = (time12h: any) => {
        const isPM = time12h.indexOf('PM') !== -1;
        let [hours, minutes] = time12h.replace(isPM ? 'PM' : 'AM', '').split(':');

        if (isPM) {
            hours = parseInt(hours, 10) + 12;
            hours = hours === 24 ? 12 : hours;
        } else {
            hours = parseInt(hours, 10);
            hours = hours === 12 ? 0 : hours;
            if (String(hours).length === 1) hours = '0' + hours;
        }

        const time = [hours, minutes].join(':');

        return time;
    }

    private onGroupSelected = (value: any) => {
        this.setState({
            selectedRadioBtn: value,
            teamsOptionSelected: value === 'teams',
            rostersOptionSelected: value === 'rosters',
            adGroupOptionSelected: value === "adgroups",
            allUsersOptionSelected: value === 'allUsers',
            selectedTeams: value === 'teams' ? this.state.selectedTeams : [],
            selectedTeamsNum: value === 'teams' ? this.state.selectedTeamsNum : 0,
            selectedRosters: value === 'rosters' ? this.state.selectedRosters : [],
            selectedRostersNum: value === 'rosters' ? this.state.selectedRostersNum : 0,
            selectedADGroups: value === 'adgroups' ? this.state.selectedADGroups : [],
            selectedADGroupsNum: value === 'adgroups' ? this.state.selectedADGroupsNum : 0,
        });
    }

    private isSaveBtnDisabled = () => {
        const teamsSelectionIsValid = (this.state.teamsOptionSelected && (this.state.selectedTeamsNum !== 0)) || (!this.state.teamsOptionSelected);
        const rostersSelectionIsValid = (this.state.rostersOptionSelected && (this.state.selectedRostersNum !== 0)) || (!this.state.rostersOptionSelected);
        const adGroupsSelectionIsValid = (this.state.adGroupOptionSelected && (this.state.selectedADGroupsNum !== 0)) || (!this.state.adGroupOptionSelected);
        const nothingSelected = (!this.state.teamsOptionSelected) && (!this.state.rostersOptionSelected) && (!this.state.adGroupOptionSelected) && (!this.state.allUsersOptionSelected);

        return (!teamsSelectionIsValid || !rostersSelectionIsValid || !adGroupsSelectionIsValid || nothingSelected || !this.state.isRepeatTimeValid || !this.state.isWeeksSelectionValid)
    }

    private isNextBtnDisabled = () => {
        const title = this.state.title;
        const btnTitle = this.state.btnTitle;
        const btnLink = this.state.btnLink;
        const btnTitle2 = this.state.btnTitle2;
        const btnLink2 = this.state.btnLink2;

        return !(title && ((btnTitle && btnLink) || (!btnTitle && !btnLink)) && ((btnTitle2 && btnLink2) || (!btnTitle2 && !btnLink2)) && (this.state.errorImageUrlMessage === "") && (this.state.errorButtonUrlMessage === ""));
    }

    private getItems = () => {
        const resultedTeams: dropdownItem[] = [];
        if (this.state.teams) {
            let remainingUserTeams = this.state.teams;
            if (this.state.selectedRadioBtn !== "allUsers") {
                remainingUserTeams = this.state.selectedRadioBtn === "teams" ? this.state.teams.filter(x => this.state.selectedTeams.findIndex(y => y.team.id === x.teamId) < 0) : this.state.teams.filter(x => this.state.selectedRosters.findIndex(y => y.team.id === x.teamId) < 0);
            }
            remainingUserTeams.forEach((element) => {
                resultedTeams.push({
                    header: element.name,
                    team: {
                        id: element.teamId
                    }
                });
            });
        }
        return resultedTeams;
    }

    private getADGroupItems = async (event: any, itemsData: any) => {

        if (!isNullOrEmpty(itemsData.searchQuery)) {
            const response = await getADGroups(itemsData.searchQuery);
            let adGroups: ADGroup[] = [];
            if (!isNullOrUndefined(response.data)) {
                adGroups = response.data
            }

            const resultedTeams: dropdownItem[] = [];
            adGroups = adGroups.filter(x => this.state.selectedADGroups.findIndex(y => y.team.id === x.id) < 0);

            adGroups.forEach((element) => {
                resultedTeams.push({
                    header: element.displayName,
                    team: {
                        id: element.id
                    }
                });
            });


            this.setState({
                resultADGroups: resultedTeams
            })
        }
    }

    private onTeamsChange = (event: any, itemsData: any) => {
        this.setState({
            selectedTeams: itemsData.value,
            selectedTeamsNum: itemsData.value.length,
            selectedRosters: [],
            selectedRostersNum: 0,
            selectedADGroups: [],
            selectedADGroupsNum: 0
        })
    }

    private onRostersChange = (event: any, itemsData: any) => {
        this.setState({
            selectedRosters: itemsData.value,
            selectedRostersNum: itemsData.value.length,
            selectedTeams: [],
            selectedTeamsNum: 0,
            selectedADGroups: [],
            selectedADGroupsNum: 0
        })
    }

    private onADGroupChange = (event: any, itemsData: any) => {
        this.setState({
            selectedADGroups: itemsData.value,
            selectedADGroupsNum: itemsData.value.length,
            selectedTeams: [],
            selectedTeamsNum: 0,
            selectedRosters: [],
            selectedRostersNum: 0
        })
    }

    private onRepeatChange = (event: any, itemsData: any) => {
        this.setState({
            repeats: itemsData.value.header,
        })
    }

    private onRepeatFrequencyChange = (event: any, itemsData: any) => {
        this.setState({
            repeatFrequency: itemsData.value.header,
        },
            () => {
                console.log(this.state.repeats + "" + this.state.repeatFrequency + !this.state.mondaySelection + !this.state.tuesdaySelection + !this.state.webnesdaySelection + !this.state.thursdaySelection + !this.state.fridaySelection + !this.state.saturdaySelection + !this.state.sundaySelection);
                console.log(repeatsValues[5] + " " + repeatFrequency[1]);
                if (this.state.repeats === repeatsValues[5] && this.state.repeatFrequency === repeatFrequency[1] && !this.state.mondaySelection && !this.state.tuesdaySelection && !this.state.webnesdaySelection && !this.state.thursdaySelection && !this.state.fridaySelection && !this.state.saturdaySelection && !this.state.sundaySelection) {
                    console.log("Entered");
                    this.setState({
                        isWeeksSelectionValid: false,
                    })
                }
                else {
                    this.setState({
                        isWeeksSelectionValid: true,
                    })
                }
            });

    }

    private onSave = () => {
        let draftMessage = this.saveMessage();

        if (this.state.exists) {
            this.editDraftMessage(draftMessage).then(() => {
                microsoftTeams.tasks.submitTask();
            });
        } else {
            this.postDraftMessage(draftMessage).then(() => {
                microsoftTeams.tasks.submitTask();
            });
        }
    }

    private onSend = () => {
        let draftMessage = this.saveMessage();
        let notificationId: string = "";

        if (this.state.exists) {
            this.editDraftMessage(draftMessage).then((response: any) => {
                notificationId = this.state.messageId;
                draftMessage.id = notificationId;
                let spanner = document.getElementsByClassName("sendingLoader");
                spanner[0].classList.remove("hiddenLoader");

                sendDraftNotification(draftMessage).then(() => {
                    microsoftTeams.tasks.submitTask();
                });
            });
        } else {
            this.postDraftMessage(draftMessage).then((response: any) => {
                notificationId = response.data;
                draftMessage.id = notificationId;
                let spanner = document.getElementsByClassName("sendingLoader");
                spanner[0].classList.remove("hiddenLoader");

                sendDraftNotification(draftMessage).then(() => {
                    microsoftTeams.tasks.submitTask();
                });

            });
        }

    }

    private saveMessage = () => {
        const selectedTeams: string[] = [];
        const selctedRosters: string[] = [];
        const selctedADGroups: string[] = [];
        this.state.selectedTeams.forEach(x => selectedTeams.push(x.team.id));
        this.state.selectedRosters.forEach(x => selctedRosters.push(x.team.id));
        this.state.selectedADGroups.forEach(x => selctedADGroups.push(x.team.id));

        let weekSelection = "";
        if (this.state.isRecurrenceEnabled) {

            if (this.state.mondaySelection) {
                weekSelection = "0/";
            }

            if (this.state.tuesdaySelection) {
                weekSelection = weekSelection + "1/";
            }

            if (this.state.webnesdaySelection) {
                weekSelection = weekSelection + "2/";
            }

            if (this.state.thursdaySelection) {
                weekSelection = weekSelection + "3/";
            }

            if (this.state.fridaySelection) {
                weekSelection = weekSelection + "4/";
            }

            if (this.state.saturdaySelection) {
                weekSelection = weekSelection + "5/";
            }
            if (this.state.sundaySelection) {
                weekSelection = weekSelection + "6/";
            }
        }

        let scheduleDate = this.state.scheduleDate;
        let repeatsFor = this.state.repeatsFor;
        let repeats = this.state.repeats;
        let repeatFrequency = this.state.repeatFrequency;
        let repeatStartDate: Date = this.state.repeatsStartDate;
        let repeatsEndDate: Date = this.state.repeatsEndDate;
        let currentDate = this.datetime(new Date(), "00:00 AM");

        if (!this.state.isRecurrenceEnabled) {
            repeatStartDate = currentDate;
            repeatsEndDate = currentDate;
            repeatsFor = 0;
            repeats = "";
            repeatFrequency = "";
        }

        if (!this.state.IsScheduled) {
            scheduleDate = currentDate;
        }

        let draftMessage: IDraftMessage = {
            id: this.state.messageId,
            title: this.state.title,
            imageLink: this.state.imageLink,
            summary: this.state.summary,
            author: this.state.author,
            buttonTitle: this.state.btnTitle,
            buttonLink: this.state.btnLink,
            teams: selectedTeams,
            rosters: selctedRosters,
            adGroups: selctedADGroups,
            allUsers: this.state.allUsersOptionSelected,
            buttonTitle2: this.state.btnTitle2,
            buttonLink2: this.state.btnLink2,
            IsScheduled: this.state.IsScheduled,
            scheduleDate: scheduleDate,
            isRecurrence: this.state.isRecurrenceEnabled,
            repeats: repeats,
            repeatFor: repeatsFor,
            repeatFrequency: repeatFrequency,
            weekSelection: weekSelection,
            repeatStartDate: repeatStartDate,
            repeatEndDate: repeatsEndDate
        };

        return draftMessage;
    }

    private editDraftMessage = async (draftMessage: IDraftMessage) => {
        try {
            return await updateDraftNotification(draftMessage);
        } catch (error) {
            return error;
        }
    }

    private postDraftMessage = async (draftMessage: IDraftMessage) => {
        try {
            return await createDraftNotification(draftMessage);
        } catch (error) {
            return error;
        }
    }

    public escFunction(event: any) {
        if (event.keyCode === 27 || (event.key === "Escape")) {
            microsoftTeams.tasks.submitTask();
        }
    }

    private onNext = (event: any) => {
        this.setState({
            page: "AudienceSelection"
        }, () => {
            this.updateCard();
        });
    }

    private onBack = (event: any) => {
        this.setState({
            page: "CardCreation"
        }, () => {
            this.updateCard();
        });
    }

    private onRepeatForChanged = (event: any) => {
        if (Number(event.target.value)) {
            this.setState({
                repeatsFor: event.target.value,
            });
        }
    }
    private onTitleChanged = (event: any) => {
        let showDefaultCard = (!event.target.value && !this.state.imageLink && !this.state.summary && !this.state.author && !this.state.btnTitle && !this.state.btnLink);
        setCardTitle(this.card, event.target.value);
        setCardImageLink(this.card, this.state.imageLink);
        setCardSummary(this.card, this.state.summary);
        setCardAuthor(this.card, this.state.author);
        setCardBtn(this.card, this.state.btnTitle, this.state.btnLink, this.state.btnTitle2, this.state.btnLink2);
        this.setState({
            title: event.target.value,
            card: this.card
        }, () => {
            if (showDefaultCard) {
                this.setDefaultCard(this.card);
            }
            this.updateCard();
        });
    }

    private onImageLinkChanged = (event: any) => {
        let url = event.target.value.toLowerCase();
        if (!((url === "") || (url.startsWith("https://") || (url.startsWith("data:image/png;base64,")) || (url.startsWith("data:image/jpeg;base64,")) || (url.startsWith("data:image/gif;base64,"))))) {
            this.setState({
                errorImageUrlMessage: "URL must start with https://"
            });
        } else {
            this.setState({
                errorImageUrlMessage: ""
            });
        }

        let showDefaultCard = (!this.state.title && !event.target.value && !this.state.summary && !this.state.author && !this.state.btnTitle && !this.state.btnLink && !this.state.btnTitle2 && !this.state.btnLink2);
        setCardTitle(this.card, this.state.title);
        setCardImageLink(this.card, event.target.value);
        setCardSummary(this.card, this.state.summary);
        setCardAuthor(this.card, this.state.author);
        setCardBtn(this.card, this.state.btnTitle, this.state.btnLink, this.state.btnTitle2, this.state.btnLink2);
        this.setState({
            imageLink: event.target.value,
            card: this.card
        }, () => {
            if (showDefaultCard) {
                this.setDefaultCard(this.card);
            }
            this.updateCard();
        });
    }

    private onSummaryChanged = (event: any) => {
        let showDefaultCard = (!this.state.title && !this.state.imageLink && !event.target.value && !this.state.author && !this.state.btnTitle && !this.state.btnLink && !this.state.btnTitle2 && !this.state.btnLink2);
        setCardTitle(this.card, this.state.title);
        setCardImageLink(this.card, this.state.imageLink);
        setCardSummary(this.card, event.target.value);
        setCardAuthor(this.card, this.state.author);
        setCardBtn(this.card, this.state.btnTitle, this.state.btnLink, this.state.btnTitle2, this.state.btnLink2);
        this.setState({
            summary: event.target.value,
            card: this.card
        }, () => {
            if (showDefaultCard) {
                this.setDefaultCard(this.card);
            }
            this.updateCard();
        });
    }

    private onAuthorChanged = (event: any) => {
        let showDefaultCard = (!this.state.title && !this.state.imageLink && !this.state.summary && !event.target.value && !this.state.btnTitle && !this.state.btnLink);
        setCardTitle(this.card, this.state.title);
        setCardImageLink(this.card, this.state.imageLink);
        setCardSummary(this.card, this.state.summary);
        setCardAuthor(this.card, event.target.value);
        setCardBtn(this.card, this.state.btnTitle, this.state.btnLink);
        this.setState({
            author: event.target.value,
            card: this.card
        }, () => {
            if (showDefaultCard) {
                this.setDefaultCard(this.card);
            }
            this.updateCard();
        });
    }

    private onBtnTitleChanged = (event: any) => {
        const showDefaultCard = (!this.state.title && !this.state.imageLink && !this.state.summary && !this.state.author && !event.target.value && !this.state.btnLink && !this.state.btnTitle2 && !this.state.btnLink2);
        setCardTitle(this.card, this.state.title);
        setCardImageLink(this.card, this.state.imageLink);
        setCardSummary(this.card, this.state.summary);
        setCardAuthor(this.card, this.state.author);
        if ((event.target.value && this.state.btnLink) || (this.state.btnTitle2 && this.state.btnLink2)) {
            setCardBtn(this.card, event.target.value, this.state.btnLink, this.state.btnTitle2, this.state.btnLink2);
            this.setState({
                btnTitle: event.target.value,
                card: this.card
            }, () => {
                if (showDefaultCard) {
                    this.setDefaultCard(this.card);
                }
                this.updateCard();
            });
        } else {
            delete this.card.actions;
            this.setState({
                btnTitle: event.target.value,
            }, () => {
                if (showDefaultCard) {
                    this.setDefaultCard(this.card);
                }
                this.updateCard();
            });
        }
    }

    private onBtnLinkChanged = (event: any) => {
        if (!(event.target.value === "" || event.target.value.toLowerCase().startsWith("https://"))) {
            this.setState({
                errorButtonUrlMessage: "URL must start with https://"
            });
        } else {
            this.setState({
                errorButtonUrlMessage: ""
            });
        }

        const showDefaultCard = (!this.state.title && !this.state.imageLink && !this.state.summary && !this.state.author && !this.state.btnTitle && !event.target.value && !this.state.btnTitle2 && !this.state.btnLink2);
        setCardTitle(this.card, this.state.title);
        setCardSummary(this.card, this.state.summary);
        setCardAuthor(this.card, this.state.author);
        setCardImageLink(this.card, this.state.imageLink);
        if (this.state.btnTitle && event.target.value) {
            setCardBtn(this.card, this.state.btnTitle, event.target.value, this.state.btnTitle2, this.state.btnLink2);
            this.setState({
                btnLink: event.target.value,
                card: this.card
            }, () => {
                if (showDefaultCard) {
                    this.setDefaultCard(this.card);
                }
                this.updateCard();
            });
        } else {
            delete this.card.actions;
            this.setState({
                btnLink: event.target.value
            }, () => {
                if (showDefaultCard) {
                    this.setDefaultCard(this.card);
                }
                this.updateCard();
            });
        }
    }

    private onBtnTitleChanged2 = (event: any) => {
        const showDefaultCard = (!this.state.title && !this.state.imageLink && !this.state.summary && !this.state.author && !this.state.btnTitle && !this.state.btnLink && !event.target.value && !this.state.btnLink2);
        setCardTitle(this.card, this.state.title);
        setCardImageLink(this.card, this.state.imageLink);
        setCardSummary(this.card, this.state.summary);
        setCardAuthor(this.card, this.state.author);
        if ((this.state.btnTitle && this.state.btnLink) || (event.target.value && this.state.btnLink2)) {
            setCardBtn(this.card, this.state.btnTitle, this.state.btnLink, event.target.value, this.state.btnLink2);
            this.setState({
                btnTitle2: event.target.value,
                card: this.card
            }, () => {
                if (showDefaultCard) {
                    this.setDefaultCard(this.card);
                }
                this.updateCard();
            });
        } else {
            delete this.card.actions;
            this.setState({
                btnTitle2: event.target.value,
            }, () => {
                if (showDefaultCard) {
                    this.setDefaultCard(this.card);
                }
                this.updateCard();
            });
        }
    }

    private onBtnLinkChanged2 = (event: any) => {
        if (!(event.target.value === "" || event.target.value.toLowerCase().startsWith("https://"))) {
            this.setState({
                errorButtonUrlMessage2: "URL must start with https://"
            });
        } else {
            this.setState({
                errorButtonUrlMessage2: ""
            });
        }

        const showDefaultCard = (!this.state.title && !this.state.imageLink && !this.state.summary && !this.state.author && !this.state.btnTitle && !this.state.btnLink2 && !this.state.btnTitle2 && !event.target.value);
        setCardTitle(this.card, this.state.title);
        setCardSummary(this.card, this.state.summary);
        setCardAuthor(this.card, this.state.author);
        setCardImageLink(this.card, this.state.imageLink);
        if ((this.state.btnTitle && this.state.btnLink2) || (this.state.btnTitle2 && event.target.value)) {
            setCardBtn(this.card, this.state.btnTitle, event.target.value, this.state.btnTitle2, event.target.value);
            this.setState({
                btnLink2: event.target.value,
                card: this.card
            }, () => {
                if (showDefaultCard) {
                    this.setDefaultCard(this.card);
                }
                this.updateCard();
            });
        } else {
            delete this.card.actions;
            this.setState({
                btnLink2: event.target.value
            }, () => {
                if (showDefaultCard) {
                    this.setDefaultCard(this.card);
                }
                this.updateCard();
            });
        }
    }

    private updateCard = () => {
        const adaptiveCard = new AdaptiveCards.AdaptiveCard();
        adaptiveCard.parse(this.state.card);
        const renderedCard = adaptiveCard.render();
        const container = document.getElementsByClassName('adaptiveCardContainer')[0].firstChild;
        if (container != null) {
            container.replaceWith(renderedCard);
        } else {
            document.getElementsByClassName('adaptiveCardContainer')[0].appendChild(renderedCard);
        }
        const primaryButtonTitle = this.state.btnTitle;
        const primaryButtonLink = this.state.btnLink;
        const secondaryButtonLink = this.state.btnLink2;
        adaptiveCard.onExecuteAction = function (action) {
            if (action.title === primaryButtonTitle) {
                window.open(primaryButtonLink, '_blank');
            }
            else {
                window.open(secondaryButtonLink, '_blank');
            }
        }
    }
}