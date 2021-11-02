import { env } from './env.js';

microsoftTeams.initialize();
microsoftTeams.getContext((context) => {

    // Build an Azure AD request that authenticate the user and ask them to
    // and to consent to any missing permissions
    const tenant_id = context['tid'];
    const client_id = env.CLIENT_ID;
    let queryParams = {
        tenant: tenant_id,
        client_id: client_id,
        response_type: "code", // We won't use the auth code, just forcing consent
        scope: env.SCOPES,
        redirect_uri: window.location.origin + "/consent-popup-end.html"
    }

    const authorizeEndpoint =
     `https://login.microsoftonline.com/${tenant_id}/oauth2/v2.0/authorize?` +
     new URLSearchParams(queryParams).toString();

    //Redirect to the Azure authorization endpoint. When that flow completes, the user will be directed to auth-end
    //Go to ClosePopup.js
    window.location.assign(authorizeEndpoint);

});