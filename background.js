/*global browser*/
/*jslint es6 */

'use strict';

let teamsUrl = 'https://teams.microsoft.com';
let pollUrl = 'https://emea.ng.msg.teams.microsoft.com/v1/users/ME/endpoints/SELF/subscriptions/0/poll';

function process_chat_link(link) {
  link = link.replace('db3pv2.ng.msg.', '');
  link = link.replace('v1/users/ME', '_#');
  link += '?ctx=chat';
  return link;
}

function process_topic_link(link, threadtopic, messageId) {
  link = link.split(';messageid', 1)[0]
  link = link.replace('db3pv2.ng.msg.', '');
  link = link.replace('v1/users/ME', '_#');
  link = link.replace('_#/conversations/', '_#/conversations/' + threadtopic + '?threadId=');
  link += '&ctx=channel';
  link += ('&messageId=' + messageId);
  return link;
}


function createNotification(resource) {
  let link = null;
  let threadtopic = resource.threadtopic || 'unknown';
  if (resource.threadtype == 'chat') {
    link = process_chat_link(resource.conversationLink);
  }
  else if (resource.threadtype == 'topic' || resource.threadtype == 'space') {
    link = process_topic_link(
      resource.conversationLink, threadtopic, resource.id
    );
    console.log(link);
  }
  else return;
  browser.notifications.create(link, {
    type: "basic",
    iconUrl: browser.extension.getURL("images/teams256.png"),
    title: resource.imdisplayname + (!threadtopic.includes(':orgid:')?' in ' + threadtopic: ''),
    message: resource.messagetype == 'Text'?resource.content:'New ' + resource.messagetype + ' message'
  });
}


function handleMessage(event_message) {
  if (event_message.resourceType && event_message.resourceType == 'NewMessage') {
    let resource = event_message.resource;
    if (resource.type != 'Message') return;
    if (!resource.content) return;
    console.log("checking if teams tab focused");
    checkFocused(teamsUrl + '/*').then((isFocused) => {
      if (!isFocused) createNotification(resource);
    });
  }
}


function checkFocused(urlPattern) {
  return new Promise(
      (resolve, reject) => {
          browser.windows.getLastFocused().then(
              (window) => {
                  if (!window.focused) resolve(false);
                  let query_info = {
                      url: [urlPattern],
                      active: true,
                      windowId: window.id
                  };
                  browser.tabs.query(query_info).then(
                      (tabs) => {resolve(tabs.length == 1);}
                  );
              }
          );
      }
  )
}


function onPoll(details) {
  let tab_id = details.tabId;
  if (tab_id == -1) { // Not originated from a browser TAB.
    return; // Skip if that's the case.
  }
  let filter = browser.webRequest.filterResponseData(details.requestId);
  let decoder = new TextDecoder("utf-8");
  let encoder = new TextEncoder();

  let string_data = '';
  filter.ondata = (event) => {
    string_data += decoder.decode(event.data, {stream: true});
  };

  filter.onstop = (event) => {
    try {
      console.log(string_data);
      if (string_data.length > 0) {
        let json_data = JSON.parse(string_data);
        if (json_data.eventMessages) {
          for (let i = 0; i < json_data.eventMessages.length; i++) {
            handleMessage(json_data.eventMessages[i]);
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

function goToTeams() {
  return new Promise(
    (resolve, reject) => {
      console.log("Navigating to teams tab.");
      browser.tabs.query({url: [teamsUrl + '/*']}).then(
        (tabs) => tabs[0]
      ).then(
        (tab) => {
          if (tab) {
            browser.windows.update(tab.windowId, {focused: true}).then(
              (window) => {
                browser.tabs.update(tab.id, {active: true}).then(resolve);
              }
            );
          } else {
            browser.tabs.create({url: teamsUrl}).then(resolve);
          }
        }
      );
    }
  )
}

function goToTeamsURL(url) {
  return new Promise(
    (resolve, reject) => {
      goToTeams().then(
        (tab) => {
          console.log(tab);
          browser.tabs.update(
            tab.id, {url: url, loadReplace: true}
          ).then(resolve)
        }
      );
    }
  )
}


browser.runtime.onInstalled.addListener(() => {
  browser.browserAction.onClicked.addListener(goToTeams);
  browser.notifications.onClicked.addListener(function(notificationId) {
    browser.notifications.clear(notificationId);
    goToTeamsURL(notificationId);
  });
  browser.webRequest.onBeforeRequest.addListener(
    onPoll,
    {urls: [pollUrl]},
    ['blocking']
  )
});
