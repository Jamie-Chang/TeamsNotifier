# TeamsNotifier
A Firefox extension to send notifications for new messages in MS Teams.

## Why do I need this?
The web app for Microsoft teams does not push desktop notifications, a strange feature to be missing in an instant messenging app. [This maybe fixed in the future](https://microsoftteams.uservoice.com/forums/555103-public/suggestions/18671746-browser-os-independent-desktop-notifications-using#{toggle_previous_statuses}) but in the mean time this extension should do the job.

## Why no chrome?
This project originally started as a chrome extension, unfortunately chrome's API is lacking some features that are essential for the extension, see https://bugs.chromium.org/p/chromium/issues/detail?id=487422.

## TODO:
* ~~Focus window when clicking on notification or icon.~~
* ~~Handle larger messages properly.~~
* ~~Show notifications for channels.~~
* ~~Only show notifications for followed channels.~~
* Handle uncommon message types: Likes, mentions etc.
