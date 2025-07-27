// Microsoft Graph API Configuration
// To make this work, you need to register an application in Azure Active Directory
// and replace the clientId with your actual application ID

const msalConfig = {
    auth: {
        clientId: "YOUR_CLIENT_ID_HERE", // Replace with your Azure AD application client ID
        authority: "https://login.microsoftonline.com/common",
        redirectUri: window.location.origin // This will be your app's URL
    },
    cache: {
        cacheLocation: "sessionStorage", // Use sessionStorage or localStorage
        storeAuthStateInCookie: false
    }
};

// Scopes required for Microsoft Graph API
const graphScopes = {
    scopes: [
        "User.Read",           // Read user profile
        "Mail.Read",           // Read user's mail
        "Calendars.Read"       // Read user's calendar
    ]
};

// Microsoft Graph API endpoints
const graphConfig = {
    graphMeEndpoint: "https://graph.microsoft.com/v1.0/me",
    graphMailEndpoint: "https://graph.microsoft.com/v1.0/me/messages?$top=5&$orderby=receivedDateTime desc",
    graphCalendarEndpoint: "https://graph.microsoft.com/v1.0/me/events?$top=5&$orderby=start/dateTime"
};

// Configuration notes:
// 1. Register your application at https://portal.azure.com/
// 2. Go to "Azure Active Directory" > "App registrations" > "New registration"
// 3. Set the redirect URI to match your application URL
// 4. Copy the "Application (client) ID" and replace YOUR_CLIENT_ID_HERE
// 5. Under "API permissions", add Microsoft Graph permissions:
//    - User.Read (delegated)
//    - Mail.Read (delegated)
//    - Calendars.Read (delegated)
// 6. Grant admin consent for the permissions if required