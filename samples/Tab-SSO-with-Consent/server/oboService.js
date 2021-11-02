import fetch from 'node-fetch';
import { INTERACTION_REQUIRED_STATUS_TEXT } from "./constants.js";

export async function getOboAccessToken(tenantId, clientSideToken) {
    
    const clientId = process.env.CLIENT_ID;
    const clientSecret = process.env.CLIENT_SECRET;
    const scopes = process.env.SCOPES;

    // Use On Behalf Of flow to exchange the client-side token for an
    // access token with the needed permissions
    const url = "https://login.microsoftonline.com/" + tenantId + "/oauth2/v2.0/token";
    const params = {
        client_id: clientId,
        client_secret: clientSecret,
        grant_type: "urn:ietf:params:oauth:grant-type:jwt-bearer",
        assertion: clientSideToken,
        requested_token_use: "on_behalf_of",
        scope: scopes
    };

    const accessTokenQueryParams = new URLSearchParams(params).toString();
    const oboResponse = await fetch(url, {
        method: "POST",
        body: accessTokenQueryParams,
        headers: {
            Accept: "application/json",
            "Content-Type": "application/x-www-form-urlencoded"
        }
    });

    const oboData = await oboResponse.json();
    if (oboResponse.status !== 200) {
        // We got an error on the OBO request. Check if it is consent required.
        if (oboData.error.toLowerCase() === 'invalid_grant' ||
            oboData.error.toLowerCase() === 'interaction_required') {
            throw(INTERACTION_REQUIRED_STATUS_TEXT);
        } else {
            console.log(`Error returned in OBO: ${JSON.stringify(oboData)}`);
            throw (`Error in OBO exchange ${oboResponse.status}: ${oboResponse.statusText}`);
        }
    }

    return oboData.access_token;
    
}