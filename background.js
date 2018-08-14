/*global browser*/

'use strict';


function process_chat_link(link) {
  link = link.replace('db3pv2.ng.msg.', '')
  link = link.replace('v1/users/ME', '_#')
  link += '?ctx=chat'
  return link
}


function handle_message(event_message) {
  if (event_message.resourceType && event_message.resourceType == 'NewMessage') {
    let resource = event_message.resource
    if (resource.type == 'Message' &&
      resource.threadtype == 'chat' &&
      resource.messagetype == 'Text'
    ){
      browser.notifications.create(process_chat_link(resource.conversationLink), {
        "type": "basic",
        "iconUrl": browser.extension.getURL("images/teams256.png"),
        "title": resource.imdisplayname + (!resource.threadtopic.includes(':orgid:')?' in ' + resource.threadtopic: ''),
        "message": resource.messagetype == 'Text'?resource.content:'...'
      })
    }
  }
  // TODO: Handle ConversationUpdate resourceType
}


function onPollRequest(details) {
  let filter = browser.webRequest.filterResponseData(details.requestId);
  let decoder = new TextDecoder("utf-8");
  filter.ondata = function(event) {
    try {
      let data = JSON.parse(decoder.decode(event.data, {stream: true}))
      if (data.eventMessages) {
        for (let i = 0; i < data.eventMessages.length; i++) {
          handle_message(data.eventMessages[i]);
        }
      }
    }
    finally {
      filter.write(event.data);
      filter.disconnect();
    }
  }
}

var teams = {
  tab: null,
  url: 'https://teams.microsoft.com',
  resetTab: function() {
    this.tab = null;
    browser.browserAction.setIcon({path: 'images/teamsgrey38.png'});
    browser.webRequest.onBeforeRequest.removeListener(onPollRequest);
  },
  setTab: function(tab) {
    this.tab = tab;
    browser.browserAction.setIcon({path: 'images/teams38.png'});
    console.log("Setting tab")
    browser.webRequest.onBeforeRequest.addListener(
      onPollRequest,
      {
        tabId: this.tab.id,
        urls: ['https://emea.ng.msg.teams.microsoft.com/v1/users/ME/endpoints/SELF/subscriptions/0/poll']
      },
      ['blocking']
    )
  }
}


function findTeams() {
  let queryInfo = {url: [teams.url + '/*']};
  browser.tabs.query(queryInfo, function(tabs) {
    console.log(tabs)
    for (let i = 0; i < tabs.length; i++) {
      let tab = tabs[i];
      teams.setTab(tab);
      return;
    }
  });
  if (teams.tab === null) teams.resetTab();
}


function goToTeams() {
  if (teams.tab !== null) {
    browser.tabs.update(teams.tab.id, {active: true});
    return;
  }
  findTeams();
  if (teams.tab === null) browser.tabs.create({url: teams.url}, function(tab) {teams.setTab(tab);});
  else browser.tabs.update(teams.tab.id, {active: true});
}

function goToTeamsChat(link) {
  goToTeams();
  browser.tabs.update(teams.tab.id, {url: link})
}

function createdCallback(tab) {
  if (teams.tab === null && tab.url.startsWith(teams.url)) {
    teams.setTab(tab);
  }
}


function updatedCallback(tab_id, change_info, tab) {
  if (!change_info.url) return;
  if (teams.tab!== null && tab.id === teams.tab.id && !change_info.url.startsWith(teams.url)) {
    // Teams tab changed URL
    teams.resetTab();
  }
  else createdCallback(tab);
}


function removedCallback(tab_id, remove_info) {
  if (teams.tab !== null && teams.tab.id == tab_id) {
    teams.resetTab();
  }
}

browser.browserAction.onClicked.addListener(goToTeams);
browser.runtime.onInstalled.addListener(function() {
  browser.tabs.onCreated.addListener(createdCallback);
  browser.tabs.onRemoved.addListener(removedCallback);
  browser.tabs.onUpdated.addListener(updatedCallback);
});
findTeams();
browser.notifications.onClicked.addListener(function(notification_id) {
  browser.notifications.clear(notification_id);
  goToTeamsChat(notification_id);
});
