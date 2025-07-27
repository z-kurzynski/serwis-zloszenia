// Microsoft Graph API Integration for Outlook Connection
// This script handles authentication and data retrieval from Microsoft Graph

let msalInstance;
let currentUser = null;

// Initialize MSAL on page load
document.addEventListener('DOMContentLoaded', function() {
    initializeMsal();
    setupEventListeners();
    checkExistingLogin();
});

// Initialize Microsoft Authentication Library
function initializeMsal() {
    try {
        msalInstance = new msal.PublicClientApplication(msalConfig);
        console.log('MSAL initialized successfully');
    } catch (error) {
        console.error('Failed to initialize MSAL:', error);
        showError('Błąd inicjalizacji systemu uwierzytelniania. Sprawdź konfigurację.');
    }
}

// Set up event listeners for buttons
function setupEventListeners() {
    document.getElementById('login-btn').addEventListener('click', signIn);
    document.getElementById('logout-btn').addEventListener('click', signOut);
    document.getElementById('retry-btn').addEventListener('click', hideError);
}

// Check if user is already logged in
async function checkExistingLogin() {
    try {
        const accounts = msalInstance.getAllAccounts();
        if (accounts.length > 0) {
            currentUser = accounts[0];
            await getTokenAndLoadData();
        }
    } catch (error) {
        console.error('Error checking existing login:', error);
    }
}

// Sign in to Microsoft Graph
async function signIn() {
    try {
        showLoginStatus('Logowanie w toku...');
        
        const loginResponse = await msalInstance.loginPopup(graphScopes);
        currentUser = loginResponse.account;
        
        await getTokenAndLoadData();
        
    } catch (error) {
        console.error('Login failed:', error);
        if (error.errorCode === 'popup_window_error' || error.errorCode === 'user_cancelled') {
            showError('Logowanie zostało anulowane. Spróbuj ponownie.');
        } else {
            showError('Błąd logowania: ' + (error.errorMessage || error.message));
        }
    }
}

// Get access token and load user data
async function getTokenAndLoadData() {
    try {
        const tokenResponse = await msalInstance.acquireTokenSilent({
            ...graphScopes,
            account: currentUser
        });
        
        await loadUserData(tokenResponse.accessToken);
        showUserSection();
        
    } catch (error) {
        console.error('Token acquisition failed:', error);
        
        // If silent token acquisition fails, try with popup
        try {
            const tokenResponse = await msalInstance.acquireTokenPopup(graphScopes);
            await loadUserData(tokenResponse.accessToken);
            showUserSection();
        } catch (popupError) {
            console.error('Popup token acquisition failed:', popupError);
            showError('Nie udało się uzyskać dostępu do danych. Spróbuj zalogować się ponownie.');
        }
    }
}

// Load user profile, emails, and calendar data
async function loadUserData(accessToken) {
    try {
        // Load user profile
        const userProfile = await callGraphAPI(graphConfig.graphMeEndpoint, accessToken);
        displayUserInfo(userProfile);
        
        // Load emails
        const emails = await callGraphAPI(graphConfig.graphMailEndpoint, accessToken);
        displayEmails(emails.value || []);
        
        // Load calendar events
        const events = await callGraphAPI(graphConfig.graphCalendarEndpoint, accessToken);
        displayCalendarEvents(events.value || []);
        
    } catch (error) {
        console.error('Failed to load user data:', error);
        showError('Błąd podczas ładowania danych z Outlook. Sprawdź uprawnienia aplikacji.');
    }
}

// Make API call to Microsoft Graph
async function callGraphAPI(endpoint, accessToken) {
    const headers = {
        'Authorization': `Bearer ${accessToken}`,
        'Content-Type': 'application/json'
    };
    
    const response = await fetch(endpoint, { headers });
    
    if (!response.ok) {
        throw new Error(`Graph API error: ${response.status} ${response.statusText}`);
    }
    
    return await response.json();
}

// Display user information
function displayUserInfo(userProfile) {
    document.getElementById('user-name').textContent = userProfile.displayName || 'Użytkownik';
    document.getElementById('user-email').textContent = userProfile.mail || userProfile.userPrincipalName || '';
}

// Display email messages
function displayEmails(emails) {
    const emailsContainer = document.getElementById('emails-list');
    
    if (emails.length === 0) {
        emailsContainer.innerHTML = '<p class="text-center">Brak najnowszych wiadomości email.</p>';
        return;
    }
    
    const emailsHtml = emails.map(email => `
        <div class="email-item">
            <h4>${escapeHtml(email.subject || 'Bez tematu')}</h4>
            <p><strong>Od:</strong> ${escapeHtml(email.from?.emailAddress?.name || 'Nieznany nadawca')}</p>
            <p><strong>Adres:</strong> ${escapeHtml(email.from?.emailAddress?.address || '')}</p>
            <div class="email-meta">
                <p><strong>Otrzymano:</strong> ${formatDateTime(email.receivedDateTime)}</p>
                <p><strong>Przeczytane:</strong> ${email.isRead ? 'Tak' : 'Nie'}</p>
            </div>
        </div>
    `).join('');
    
    emailsContainer.innerHTML = emailsHtml;
}

// Display calendar events
function displayCalendarEvents(events) {
    const eventsContainer = document.getElementById('calendar-events');
    
    if (events.length === 0) {
        eventsContainer.innerHTML = '<p class="text-center">Brak nadchodzących spotkań.</p>';
        return;
    }
    
    const eventsHtml = events.map(event => `
        <div class="calendar-item">
            <h4>${escapeHtml(event.subject || 'Bez tematu')}</h4>
            <p><strong>Organizator:</strong> ${escapeHtml(event.organizer?.emailAddress?.name || 'Nieznany')}</p>
            <div class="calendar-meta">
                <p><strong>Start:</strong> ${formatDateTime(event.start?.dateTime)}</p>
                <p><strong>Koniec:</strong> ${formatDateTime(event.end?.dateTime)}</p>
                <p><strong>Lokalizacja:</strong> ${escapeHtml(event.location?.displayName || 'Brak lokalizacji')}</p>
            </div>
        </div>
    `).join('');
    
    eventsContainer.innerHTML = eventsHtml;
}

// Sign out
async function signOut() {
    try {
        await msalInstance.logoutPopup();
        currentUser = null;
        showLoginSection();
    } catch (error) {
        console.error('Logout failed:', error);
        // Even if logout fails, show login section
        showLoginSection();
    }
}

// UI Helper Functions
function showLoginSection() {
    document.getElementById('login-section').classList.remove('hidden');
    document.getElementById('user-section').classList.add('hidden');
    document.getElementById('error-section').classList.add('hidden');
    document.getElementById('login-status').textContent = '';
}

function showUserSection() {
    document.getElementById('login-section').classList.add('hidden');
    document.getElementById('user-section').classList.remove('hidden');
    document.getElementById('error-section').classList.add('hidden');
    document.getElementById('user-section').classList.add('fade-in');
}

function showError(message) {
    document.getElementById('error-message').textContent = message;
    document.getElementById('login-section').classList.add('hidden');
    document.getElementById('user-section').classList.add('hidden');
    document.getElementById('error-section').classList.remove('hidden');
}

function hideError() {
    showLoginSection();
}

function showLoginStatus(message) {
    document.getElementById('login-status').textContent = message;
}

// Utility Functions
function escapeHtml(text) {
    if (!text) return '';
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
}

function formatDateTime(dateTimeString) {
    if (!dateTimeString) return 'Nieznana data';
    
    try {
        const date = new Date(dateTimeString);
        return date.toLocaleString('pl-PL', {
            year: 'numeric',
            month: 'long',
            day: 'numeric',
            hour: '2-digit',
            minute: '2-digit'
        });
    } catch (error) {
        return 'Nieprawidłowa data';
    }
}

// Handle configuration errors
window.addEventListener('load', function() {
    if (msalConfig.auth.clientId === 'YOUR_CLIENT_ID_HERE') {
        showError('Aplikacja nie jest skonfigurowana. Zaktualizuj config.js z prawidłowym Client ID z Azure AD.');
    }
});