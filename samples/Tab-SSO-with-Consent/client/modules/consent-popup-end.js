microsoftTeams.initialize();

microsoftTeams.getContext((context) => {
    // Just close the popup, there's nothing to send back to the tab
    microsoftTeams.authentication.notifySuccess();
});