import { env } from './env.js';

// Get a client side token from Teams
async function getClientSideToken() {

    microsoftTeams.initialize();
    return new Promise((resolve, reject) => {
        microsoftTeams.authentication.getAuthToken({
            successCallback: (result) => { resolve(result); },
            failureCallback: (error) => { reject(error); }
        });
    });
}

// Get the user profile from our web service
async function getUserProfile(clientSideToken) {

    if (!clientSideToken) {
        throw ("Error: No client side token provided in getUserProfile()");
    }

    // Get Teams context, converting callback to a promise so we can await it
    const context = await (() => {
        return new Promise((resolve) => {
            microsoftTeams.getContext(context => resolve(context));
        })
    })();

    // Request the user profile from our web service
    const response = await fetch('/userProfile', {
        method: 'post',
        headers: {
            'Content-Type': 'application/json'
        },
        body: JSON.stringify({
            'tenantId': context.tid,
            'clientSideToken': clientSideToken
        }),
        cache: 'default'
    });
    if (response.ok) {
        const userProfile = await response.json();
        return userProfile;
    } else {
        const error = await response.json();
        throw (error);
    }
}

// Display the consent pop-up if needed
async function showConsentPopup() {
    await microsoftTeams.authentication.authenticate({
        url: window.location.origin + "/consent-popup-start.html",
        width: 600,
        height: 535,
        successCallback: (() => {
            console.log('Got success callback');
            displayUI();
        })
    });
}

// Render the page on load or after a consent
async function displayUI() {

    const displayElement = document.getElementById('content');
    try {
        const clientSideToken = await getClientSideToken();
        const userProfile = await getUserProfile(clientSideToken);
        displayElement.innerHTML = `
            <h1>Hello ${userProfile.displayName}</h1>
            <h3>Profile Information</h3>
            <p>Mail: ${userProfile.mail}<br />
            Job Title: ${userProfile.jobTitle}<br />
            Preferred language: ${userProfile.preferredLanguage}</p>
        `;
    }
    catch (error) {
        if (error.statusText === env.INTERACTION_REQUIRED_STATUS_TEXT) {
            // If here, we had an error that indicates that Azure AD wants
            // to interact with the user, so show the consent popup
            displayElement.innerText = '';
            const button = document.createElement('button');
            button.innerText = 'Consent required';
            button.onclick = showConsentPopup;
            displayElement.appendChild(button);
        } else {
            // If here, we had some other error
            displayElement.innerText = `Error: ${JSON.stringify(error)}`;
        }
    }
}

displayUI();