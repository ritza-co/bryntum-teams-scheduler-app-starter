const userTimeZone = Intl.DateTimeFormat().resolvedOptions().timeZone;

// Create an authentication provider
const authProvider = {
    getAccessToken: async () => {
        // Call getToken in auth.js
        return await getToken();
    }
};
// Initialize the Graph client
const graphClient = MicrosoftGraph.Client.initWithMiddleware({ authProvider });


async function getMembers() {
    ensureScope("TeamMember.Read.All");
    return await graphClient
    .api('/teams/b35b8ba3-97e5-4f2e-803f-4926ac37a5ac/members')
    .get();
}

async function getAllShifts() {
    ensureScope("Schedule.Read.All");
    return await graphClient
    .api('/teams/b35b8ba3-97e5-4f2e-803f-4926ac37a5ac/schedule/shifts')
    .header("Prefer", `outlook.timezone="${userTimeZone}"`)
    .get();
}

async function createShift(name, start, end, userId, color) {
    ensureScope("Schedule.ReadWrite.All");
    return await graphClient
    .api('/teams/b35b8ba3-97e5-4f2e-803f-4926ac37a5ac/schedule/shifts')
    .post({
        "userId": userId,
        "sharedShift": {
            "displayName": name,
            "startDateTime": start,
            "endDateTime": end,
            "theme": color
        }
    });
}

async function updateShift(id, userId, name, start, end, color) {
    ensureScope("Schedule.ReadWrite.All");
    return await graphClient
    .api(`/teams/b35b8ba3-97e5-4f2e-803f-4926ac37a5ac/schedule/shifts/${id}`)
    .put({
        "userId": userId,
        "sharedShift": {
            "displayName": name,
            "startDateTime": start,
            "endDateTime": end,
            "theme": color
        }
    });
}

async function deleteShift(id) {
    ensureScope("Schedule.ReadWrite.All");
    return await graphClient
    .api(`/teams/b35b8ba3-97e5-4f2e-803f-4926ac37a5ac/schedule/shifts/${id}`)
    .delete();
}