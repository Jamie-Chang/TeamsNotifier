/*global browser*/
/*jslint es6 */

'use strict';

let teamsUrl = 'https://teams.microsoft.com';
let pollUrl = 'https://emea.ng.msg.teams.microsoft.com/v1/users/ME/endpoints/SELF/subscriptions/0/poll';
let messageManager = new Conditions();

function stripTags(html, defaultMsg) {
  let doc = new DOMParser().parseFromString(html, 'text/html');
  return doc.body.textContent||defaultMsg;
}

function getChatLink(conversationLink) {
  return conversationLink.replace(
    'db3pv2.ng.msg.', ''
  ).replace(
    'v1/users/ME', '_#'
  ) + '?ctx=chat';
}

function getTopicLink(conversationLink, threadtopic, messageId) {
  return conversationLink.split(
    ';messageid', 1
  )[0].replace(
    'db3pv2.ng.msg.', ''
  ).replace(
    'v1/users/ME', '_#'
  ).replace(
    '_#/conversations/', '_#/conversations/' + threadtopic + '?threadId='
  ) + '&ctx=channel&messageId=' + messageId;
}

function getLink(resource) {
  resource.threadtopic = resource.threadtopic || 'unknown';
  if (resource.threadtype == 'chat') {
    return getChatLink(resource.conversationLink);
  }
  if (resource.threadtype == 'topic' || resource.threadtype == 'space') {
    return getTopicLink(
      resource.conversationLink, resource.threadtopic, resource.id
    );
  }
  return null;
}


function createNotification(link, resource) {
  console.log("Creating notification:");
  let notification_props = {
    type: 'basic',
    iconUrl: browser.extension.getURL("images/teams256.png")
  };

  notification_props.title = ((resource) => {
    // Message from unnamed chat.
    if (resource.threadtopic.includes(':orgid:')) return resource.imdisplayname;

    // Reply from channel.
    if (resource.threadtype == 'space') return resource.imdisplayname + ' reply';

    // Named group chat or channel chat.
    return resource.imdisplayname + ' in ' + resource.threadtopic;
  }) (resource);

  notification_props.message = ((resource) => {
    let default_msg = 'New message';
    switch(resource.messagetype) {
      case 'Text':
        return resource.content;
      case 'RichText/Html':
        return stripTags(resource.content, default_msg);
      default:
        return default_msg;
    }
  }) (resource);

  console.log(notification_props);
  return browser.notifications.create(link, notification_props);
}


function handleMessage(event_message) {
  let fireNotification = (link, resource) => checkFocused(teamsUrl + '/*').then(
    (isFocused) => {
      if (!isFocused) {
        createNotification(link, resource).then(console.log);
      }
   }
  );

  if (event_message.resourceType && event_message.resourceType == 'NewMessage') {
    let resource = event_message.resource;
    if (resource.type != 'Message') {
      // Not a message or no content
      return;
    }
    let link = getLink(resource);
    if (link && resource.content) {
      resource.threadtopic = resource.threadtopic || 'unknown';
      if (resource.threadtype == 'topic') {
        messageManager.wait(resource.id, 2000).then(
          (message) => {fireNotification(link, resource)},
          console.log
        )
      }
      else {
        fireNotification(link, resource);
      }
    }
    else if (resource.properties && resource.properties.activity) {
      let activity = resource.properties.activity;
      if (activity.activityType == 'follow' &&
      activity.activitySubtype == 'channelNewMessage') {
        messageManager.notify(activity.sourceMessageId).then(
          console.log,
          console.log
        )
      }
    }
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

function findTeams() {
  return new Promise(
    (resolve, reject) => {
      console.log("Navigating to teams tab.");
      browser.tabs.query({url: [teamsUrl + '/*']}).then(
        (tabs) => tabs[0]
      ).then(
        (tab) => {
          if (tab) {
            resolve(tab);
          } else {
            browser.tabs.create({url: teamsUrl, active:false}).then(resolve);
          }
        }
      );
    }
  );
}

function highlightTab(tab) {
  return new Promise(
    (resolve, reject) => {
      console.log(tab);
      browser.windows.update(tab.windowId, {focused: true}).then(
        (window) => {
          browser.tabs.update(tab.id, {active: true}).then(resolve);
        }
      );
    }
  )
}


function goToTeamsURL(url) {
  return new Promise(
    (resolve, reject) => {
      findTeams().then(
        (tab) => {
          console.log(tab);
          browser.tabs.update(
            tab.id, {url: url, loadReplace: true}
          ).then(highlightTab).then(resolve);
        }
      );
    }
  );
}


browser.runtime.onInstalled.addListener(() => {
  browser.browserAction.onClicked.addListener(() => findTeams().then(highlightTab));
  browser.notifications.onClicked.addListener((notificationId) => {
    browser.notifications.clear(notificationId);
    goToTeamsURL(notificationId);
  });
  browser.webRequest.onBeforeRequest.addListener(
    onPoll,
    {urls: [pollUrl]},
    ['blocking']
  )
});
