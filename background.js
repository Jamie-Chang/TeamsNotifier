/*global browser*/
/*jslint es6 */

'use strict';


function process_chat_link(link) {
  link = link.replace('db3pv2.ng.msg.', '');
  link = link.replace('v1/users/ME', '_#');
  link += '?ctx=chat';
  return link;
}

function process_topic_link(link, threadtopic, messageId) {
  link = link.replace('db3pv2.ng.msg.', '');
  link = link.replace('v1/users/ME', '_#');
  link = link.replace('_#/conversations/', '_#/conversations/' + threadtopic + '?threadId=');
  link += '&ctx=channel';
  link += ('&messageId=' + messageId);
  return link;
}


function handle_message(event_message) {
  if (event_message.resourceType && event_message.resourceType == 'NewMessage') {
    let resource = event_message.resource;
    if (resource.type != 'Message') return;
    let link = null;
    if (resource.threadtype == 'chat') link = process_chat_link(resource.conversationLink);
    else if (resource.threadtype == 'topic') {
      link = process_topic_link(
        resource.conversationLink, resource.threadtopic, resource.id);
    }
    else return;
    browser.notifications.create(link, {
      type: "basic",
      iconUrl: browser.extension.getURL("images/teams256.png"),
      title: resource.imdisplayname + (!resource.threadtopic.includes(':orgid:')?' in ' + resource.threadtopic: ''),
      message: resource.messagetype == 'Text'?resource.content:'New ' + resource.messagetype + ' message'
    });
  }
}


function onPollRequest(details) {
  let filter = browser.webRequest.filterResponseData(details.requestId);
  let decoder = new TextDecoder("utf-8");
  let encoder = new TextEncoder();

  let string_data = '';
  filter.ondata = function(event) {
    string_data += decoder.decode(event.data, {stream: true});
  };

  filter.onstop = function(event) {
    try {
      console.log(string_data);
      if (string_data.length > 0) {
        let json_data = JSON.parse(string_data);
        if (json_data.eventMessages) {
          for (let i = 0; i < json_data.eventMessages.length; i++) {
            handle_message(json_data.eventMessages[i]);
          }
        }
      }
    }
    finally {
      filter.write(encoder.encode(string_data));
      filter.close();
    }
  };
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
  },
  selectTab: function() {
    browser.windows.update(this.tab.windowId, {focused: true})
    browser.tabs.update(this.tab.id, {active: true});
  }
}


function findTeams(skipTabId) {
  skipTabId =  (arguments.length == 1?skipTabId:null)
  let queryInfo = {url: [teams.url + '/*']};
  browser.tabs.query(queryInfo, function(tabs) {
    console.log(tabs)
    for (let i = 0; i < tabs.length; i++) {
      if (tabs[i].id != skipTabId) {
        teams.setTab(tabs[i]);
        return;
      }
    }
  });
  if (teams.tab === null) teams.resetTab();
}


function goToTeams() {
  if (teams.tab !== null) {
    teams.selectTab()
    return;
  }
  findTeams();
  if (teams.tab === null) browser.tabs.create({url: teams.url}, function(tab) {teams.setTab(tab);});
  else teams.selectTab();
}

function goToTeamsChat(link) {
  browser.tabs.update(teams.tab.id, {url: link})
  goToTeams();
}

function createdCallback(tab) {
  if (teams.tab === null && tab.url.startsWith(teams.url)) {
    teams.setTab(tab);
  }
}


function updatedCallback(tab_id, change_info, tab) {
  if (teams.tab !== null && tab.id === teams.tab.id) {
    // Teams tab changed URL
    teams.tab = tab;
    if (change_info.url && !change_info.url.startsWith(teams.url)) {
      // Teams tab URL changed
      teams.resetTab();
      findTeams();
    }
  } else createdCallback(tab);
}


function removedCallback(tab_id, remove_info) {
  if (teams.tab !== null && teams.tab.id == tab_id) {
    teams.resetTab();
    findTeams(tab_id);
  }
}

browser.browserAction.onClicked.addListener(goToTeams);
browser.runtime.onInstalled.addListener(function() {
  browser.tabs.onCreated.addListener(createdCallback);
  browser.tabs.onRemoved.addListener(removedCallback);
  browser.tabs.onUpdated.addListener(updatedCallback);
  findTeams();
});
browser.notifications.onClicked.addListener(function(notification_id) {
  browser.notifications.clear(notification_id);
  goToTeamsChat(notification_id);
});
