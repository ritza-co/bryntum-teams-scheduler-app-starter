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


export { getMembers, getSchedule, getAllShifts };