import { getToken, ensureScope } from './auth.js';
import { Client } from '@microsoft/microsoft-graph-client';

const userTimeZone = Intl.DateTimeFormat().resolvedOptions().timeZone;

// Create an authentication provider
const authProvider = {
    getAccessToken : async() => {
        // Call getToken in auth.js
        return await getToken();
    }
};

// Initialize the Graph client
const graphClient = Client.initWithMiddleware({ authProvider });

async function getMembers() {
    ensureScope('TeamMember.Read.All');
    return await graphClient
        .api(`/teams/${import.meta.env.VITE_MICROSOFT_TEAMS_ID}/members`)
        .get();
}

async function getSchedule() {
    ensureScope('Schedule.Read.All');
    return await graphClient
        .api(`/teams/${import.meta.env.VITE_MICROSOFT_TEAMS_ID}/schedule`)
        .get();
}

async function getAllShifts() {
    ensureScope('Schedule.Read.All');
    return await graphClient
        .api(`/teams/${import.meta.env.VITE_MICROSOFT_TEAMS_ID}/schedule/shifts`)
        .header('Prefer', `outlook.timezone="${userTimeZone}"`)
        .get();
}

async function createShift(name, start, end, userId, color) {
    ensureScope('Schedule.ReadWrite.All');
    return await graphClient
        .api(`/teams/${import.meta.env.VITE_MICROSOFT_TEAMS_ID}/schedule/shifts`)
        .post({
            'userId'      : userId,
            'sharedShift' : {
                'displayName'   : name,
                'startDateTime' : start,
                'endDateTime'   : end,
                'theme'         : color
            }
        });
}

async function updateShift(id, userId, name, start, end, color) {
    ensureScope('Schedule.ReadWrite.All');
    return await graphClient
        .api(`/teams/${import.meta.env.VITE_MICROSOFT_TEAMS_ID}/schedule/shifts/${id}`)
        .put({
            'userId'      : userId,
            'sharedShift' : {
                'displayName'   : name,
                'startDateTime' : start,
                'endDateTime'   : end,
                'theme'         : color
            }
        });
}

async function deleteShift(id) {
    ensureScope('Schedule.ReadWrite.All');
    return await graphClient
        .api(`/teams/${import.meta.env.VITE_MICROSOFT_TEAMS_ID}/schedule/shifts/${id}`)
        .delete();
}

export { getMembers, createShift, updateShift, deleteShift, getSchedule, getAllShifts };