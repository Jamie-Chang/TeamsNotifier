/*global browser*/
/*jslint es6 */

'use strict';

let teamsUrl = 'https://teams.microsoft.com';
let pollUrl = 'https://*.teams.microsoft.com/*/poll';
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


async function createNotification(link, resource) {
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

    // Mention in channel chat
    if (resource.properties.mentions) {
      return resource.imdisplayname + ' mentioned you in ' + resource.threadtopic;
    }

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
  return await browser.notifications.create(link, notification_props);
}


async function handleMessage(event_message) {
  async function fireNotification(link, resource) {
    if (! await checkFocused(teamsUrl + '/*')) {
      console.log(await createNotification(link, resource));
    }
  }

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
        try {
          await messageManager.wait(resource.id, 2000)
        }
        catch (e) {
          console.log(e)
        }
        await fireNotification(link, resource);
      }
      else {
        await fireNotification(link, resource);
      }
    }
    else if (resource.properties && resource.properties.activity) {
      let activity = resource.properties.activity;
      if (
        ( activity.activityType == 'follow' &&
          activity.activitySubtype == 'channelNewMessage') ||
        ( activity.activityType == 'mention' &&
        activity.activitySubtype == 'person')
      ) {
        try {
          message = await messageManager.notify(activity.sourceMessageId)
          console.log(message)
        }
        catch (e) {
          console.error(e)
        }
      }
    }
  }
}


async function checkFocused(urlPattern) {
  let window = await browser.windows.getLastFocused();
  if (!window.focused) return false;
  let query_info = {
    url: [urlPattern],
    active: true,
    windowId: window.id
  };
  let tabs = await browser.tabs.query(query_info);
  return tabs.length == 1;
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

async function findTeams() {
  console.log("Navigating to teams tab.");
  let tabs = await browser.tabs.query({url: [teamsUrl + '/*']})
  if (tabs[0]) {
    return tabs[0];
  } else {
    return await browser.tabs.create({url: teamsUrl, active:false});
  }
}

async function highlightTab(tab) {
  console.log(tab);
  await browser.windows.update(tab.windowId, {focused: true})
  return await browser.tabs.update(tab.id, {active: true})
}


async function goToTeamsURL(url) {
  let tab = await findTeams()
  console.log(tab);
  tab = await browser.tabs.update(tab.id, {url: url, loadReplace: true});
  return await highlightTab(tab);
}


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
