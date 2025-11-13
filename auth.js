import * as msal from '@azure/msal-browser';

const msalConfig = {
    auth : {
        clientId    : import.meta.env.VITE_MICROSOFT_ENTRA_APP_ID,
        authority   : `https://login.microsoftonline.com/${import.meta.env.VITE_MICROSOFT_ENTRA_DIRECTORY_ID}`,
        redirectUri : 'http://localhost:5173'
    }
};

const msalRequest = { scopes : [] };
function ensureScope(scope) {
    if (
        !msalRequest.scopes.some((s) => s.toLowerCase() === scope.toLowerCase())
    ) {
        msalRequest.scopes.push(scope);
    }
}

// Initialize MSAL client
const msalClient = await msal.PublicClientApplication.createPublicClientApplication(msalConfig);

// Log the user in
async function signIn() {
    const authResult = await msalClient.loginPopup(msalRequest);
    sessionStorage.setItem('msalAccount', authResult.account.username);
}

async function getToken() {
    const account = sessionStorage.getItem('msalAccount');
    if (!account) {
        throw new Error(
            'User info cleared from session. Please sign out and sign in again.'
        );
    }
    try {
    // First, attempt to get the token silently
        const silentRequest = {
            scopes  : msalRequest.scopes,
            account : msalClient.getAccountByUsername(account)
        };

        const silentResult = await msalClient.acquireTokenSilent(silentRequest);
        return silentResult.accessToken;
    }
    catch (silentError) {
    // If silent requests fails with InteractionRequiredAuthError,
    // attempt to get the token interactively
        if (silentError instanceof msal.InteractionRequiredAuthError) {
            const interactiveResult = await msalClient.acquireTokenPopup(msalRequest);
            return interactiveResult.accessToken;
        }
        else {
            throw silentError;
        }
    }
}

export { signIn, getToken, ensureScope };