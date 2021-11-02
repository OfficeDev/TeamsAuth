import fetch from 'node-fetch';

export async function getUserProfile(accessToken) {

    const graphResponse = await fetch("https://graph.microsoft.com/v1.0/me",
        {
            method: 'GET',
            headers: {
                "accept": "application/json",
                "authorization": `bearer ${accessToken}`
            },
            mode: 'cors',
            cache: 'default'
        });
    if (!graphResponse.ok) {
        throw (`Error ${graphResponse.status} calling Microsoft Graph: ${graphResponse.statusText}`);
    }
    const profile = await graphResponse.json();
    return profile;
}

