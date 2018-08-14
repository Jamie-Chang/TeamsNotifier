/*global browser*/
/*jslint es6 */

'use strict';


function process_chat_link(link) {
  link = link.replace('db3pv2.ng.msg.', '');
  link = link.replace('v1/users/ME', '_#');
  link += '?ctx=chat';
  return link;
}


function handle_message(event_message) {
  if (event_message.resourceType && event_message.resourceType == 'NewMessage') {
    let resource = event_message.resource;
    if (resource.type == 'Message' &&
      resource.threadtype == 'chat' &&
      resource.messagetype == 'Text'
    ){
      browser.notifications.create(process_chat_link(resource.conversationLink), {
        "type": "basic",
        "iconUrl": browser.extension.getURL("images/teams256.png"),
        "title": resource.imdisplayname + (!resource.threadtopic.includes(':orgid:')?' in ' + resource.threadtopic: ''),
        "message": resource.messagetype == 'Text'?resource.content:'...'
      });
    }
  }
  // TODO: Handle ConversationUpdate resourceType
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
    console.log("Setting tab");
    browser.webRequest.onBeforeRequest.addListener(
      onPollRequest,
      {
        tabId: this.tab.id,
        urls: ['https://emea.ng.msg.teams.microsoft.com/v1/users/ME/endpoints/SELF/subscriptions/0/poll']
      },
      ['blocking']
    );
  },
  selectTab: function() {
    browser.windows.update(this.tab.windowId, {focused: true});
    browser.tabs.update(this.tab.id, {active: true});
  },
  findTab: function() {
    let query_info = {url: [this.url + '/*']};
    let self = this;
    browser.tabs.query(query_info, function(tabs) {
      for (let i = 0; i < tabs.length; i++) {
        let tab = tabs[i];
        self.setTab(tab);
        return;
      }
    });
    if (this.tab === null) this.resetTab();
  },
  goTab: function() {
    if (this.tab !== null) {
      this.selectTab()
      return;
    }
    this.findTab();
    if (this.tab === null) {
      let self = this;
      browser.tabs.create(
        {url: this.url},
        function(tab) {self.setTab(tab);}
      );
    }
    else this.selectTab();
  },
  goChat: function(link) {
    this.goTab();
    browser.tabs.update(teams.tab.id, {url: link});
  }
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

browser.browserAction.onClicked.addListener(function() {teams.goTab();});
browser.runtime.onInstalled.addListener(function() {
  browser.tabs.onCreated.addListener(createdCallback);
  browser.tabs.onRemoved.addListener(removedCallback);
  browser.tabs.onUpdated.addListener(updatedCallback);
});
teams.findTab();
browser.notifications.onClicked.addListener(function(notification_id) {
  browser.notifications.clear(notification_id);
  teams.goChat(notification_id);
});
