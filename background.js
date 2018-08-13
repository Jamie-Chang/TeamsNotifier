/*global browser*/

'use strict';


function doStuff(details) {
  console.log(details);
}

var teams = {
  tab: null,
  url: 'https://teams.microsoft.com',
  resetTab: function() {
    this.tab = null;
    browser.browserAction.setIcon({path: 'images/teamsgrey38.png'});
    browser.webRequest.onBeforeRequest.removeListener(doStuff);
  },
  setTab: function(tab) {
    this.tab = tab;
    browser.browserAction.setIcon({path: 'images/teams38.png'});
    browser.webRequest.onBeforeRequest.addListener(
      doStuff,
      {tabId: this.tab.id, urls: ['https://emea.ng.msg.teams.microsoft.com/v1/users/ME/endpoints/SELF/subscriptions/0/poll']}
    )
  }
}


function findTeams() {
  console.log("Finding tabs ...")
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
    browser.tabs.update(teams.tab.id, {highlighted: true});
    return;
  }
  findTeams();
  if (teams.tab === null) browser.tabs.create({url: teams.url}, function(tab) {teams.setTab(tab);});
  else browser.tabs.update(teams.tab.id, {highlighted: true});
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


browser.runtime.onInstalled.addListener(function() {
  browser.tabs.onCreated.addListener(createdCallback);
  browser.tabs.onRemoved.addListener(removedCallback);
  browser.tabs.onUpdated.addListener(updatedCallback);
});
browser.browserAction.onClicked.addListener(function(tab) {goToTeams();});
findTeams();
