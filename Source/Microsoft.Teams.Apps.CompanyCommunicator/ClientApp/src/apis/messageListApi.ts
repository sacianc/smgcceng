import axios from './axiosJWTDecorator';
import { getBaseUrl } from '../configVariables';

let baseAxiosUrl = getBaseUrl() + '/api';

export const getSentNotifications = async (): Promise<any> => {
    let url = baseAxiosUrl + "/sentnotifications";
    return await axios.get(url);
}

export const getDraftNotifications = async (): Promise<any> => {
    let url = baseAxiosUrl + "/draftnotifications";
    return await axios.get(url);
}

export const getScheduledNotifications = async (): Promise<any> => {
    let url = baseAxiosUrl + "/sentnotifications/scheduled";
    return await axios.get(url);
}

export const getSentNotification = async (id: number): Promise<any> => {
    let url = baseAxiosUrl + "/sentnotifications/" + id;
    return await axios.get(url);
}

    export const getScheduleNotification = async (id: number): Promise<any> => {
        let url = baseAxiosUrl + "/sentnotifications/schedule/" + id;
        return await axios.get(url);
}

export const getDraftNotification = async (id: number): Promise<any> => {
    let url = baseAxiosUrl + "/draftnotifications/" + id;
    return await axios.get(url);
}

export const deleteDraftNotification = async (id: number): Promise<any> => {
    let url = baseAxiosUrl + "/draftnotifications/" + id;
    return await axios.delete(url);
}

export const duplicateDraftNotification = async (id: number): Promise<any> => {
    let url = baseAxiosUrl + "/draftnotifications/duplicates/" + id;
    return await axios.post(url);
}

export const sendDraftNotification = async (payload: {}): Promise<any> => {
    let url = baseAxiosUrl + "/sentnotifications";
    return await axios.post(url, payload);
}

export const updateDraftNotification = async (payload: {}): Promise<any> => {
    let url = baseAxiosUrl + "/draftnotifications";
    return await axios.put(url, payload);
}

export const createDraftNotification = async (payload: {}): Promise<any> => {
    let url = baseAxiosUrl + "/draftnotifications";
    return await axios.post(url, payload);
}

export const getTeams = async (): Promise<any> => {
    let url = baseAxiosUrl + "/teamdata";
    return await axios.get(url);
}

export const getADGroupList = async (payload: {}): Promise<any> => {
    let url = baseAxiosUrl + "/draftnotifications/getADGroupsList";
    return await axios.post(url, payload);
}

export const getConsentSummaries = async (id: number): Promise<any> => {
    let url = baseAxiosUrl + "/draftnotifications/consentSummaries/" + id;
    return await axios.get(url);
}

export const sendPreview = async (payload: {}): Promise<any> => {
    let url = baseAxiosUrl + "/draftNotifications/previews";
    return await axios.post(url, payload);
}

export const getADGroups = async (searchQuery: string): Promise<any> => {
    let url = baseAxiosUrl + "/draftnotifications/getADGroups/" + searchQuery;
    return await axios.get(url);
}

export const getAuthenticationConsentMetadata = async (windowLocationOriginDomain: string, login_hint: string): Promise<any> => {
    let url = `${baseAxiosUrl}/authenticationMetadata/consentUrl?windowLocationOriginDomain=${windowLocationOriginDomain}&loginhint=${login_hint}`;
    return await axios.get(url, undefined, false);
}
