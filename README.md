# Serwis Zgłoszeń - Integracja z Microsoft Outlook

Aplikacja webowa umożliwiająca połączenie z Microsoft Outlook przy użyciu Microsoft Graph API.

## Funkcjonalności

- Uwierzytelnianie z Microsoft Account/Azure AD
- Wyświetlanie profilu użytkownika
- Przeglądanie najnowszych wiadomości email
- Wyświetlanie nadchodzących spotkań z kalendarza

## Konfiguracja

### 1. Rejestracja aplikacji w Azure Active Directory

1. Przejdź do [Azure Portal](https://portal.azure.com/)
2. Wybierz "Azure Active Directory" > "App registrations" > "New registration"
3. Wypełnij formularz:
   - **Name**: Serwis Zgłoszeń
   - **Supported account types**: Accounts in any organizational directory and personal Microsoft accounts
   - **Redirect URI**: Web - `https://your-app-url.com` (zastąp swoim URL)
4. Skopiuj "Application (client) ID" z strony overview

### 2. Konfiguracja uprawnień API

1. W swojej aplikacji Azure AD, przejdź do "API permissions"
2. Dodaj następujące uprawnienia Microsoft Graph (Delegated):
   - `User.Read` - Odczyt profilu użytkownika
   - `Mail.Read` - Odczyt wiadomości email
   - `Calendars.Read` - Odczyt kalendarza
3. Kliknij "Grant admin consent" jeśli wymagane

### 3. Aktualizacja konfiguracji

Edytuj plik `config.js` i zastąp `YOUR_CLIENT_ID_HERE` swoim Client ID z Azure AD:

```javascript
const msalConfig = {
    auth: {
        clientId: "twoj-client-id-tutaj",
        // ...
    }
    // ...
};
```

## Wdrożenie

Aplikacja jest skonfigurowana do automatycznego wdrożenia na Azure Static Web Apps poprzez GitHub Actions.

## Rozwój lokalny

1. Sklonuj repozytorium
2. Skonfiguruj Azure AD (jak powyżej)
3. Zaktualizuj config.js
4. Otwórz index.html w przeglądarce lub uruchom lokalny serwer HTTP

## Bezpieczeństwo

- Aplikacja używa OAuth 2.0 z PKCE dla bezpiecznego uwierzytelniania
- Tokeny są przechowywane w sessionStorage
- Wszystkie połączenia używają HTTPS
- Aplikacja wymaga tylko uprawnień do odczytu

## Wymagania

- Nowoczesna przeglądarka internetowa z obsługą JavaScript ES6+
- Konto Microsoft lub Azure AD
- Połączenie internetowe

## Wsparcie

W przypadku problemów z konfiguracją lub użytkowaniem, sprawdź:
- [Dokumentacja Microsoft Graph](https://docs.microsoft.com/graph/)
- [MSAL.js Documentation](https://docs.microsoft.com/azure/active-directory/develop/msal-js-initializing-client-applications)