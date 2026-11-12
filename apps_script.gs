// Bond Voyage – collects results into your Google Sheet (NEW VERSION – replace the old script with this one).
// Setup: open your Google Sheet -> Extensions -> Apps Script -> delete everything, paste this -> Save
//        -> Deploy -> New deployment -> type: Web app -> Execute as: Me, Who has access: Anyone -> Deploy
//        -> copy the Web app URL (ends in /exec) and paste it into settings.js on GitHub.
// If you edit this script later: Deploy -> Manage deployments -> pencil -> Version: New version -> Deploy.
function doPost(e) {
  var d = JSON.parse(e.postData.contents);
  var ss = SpreadsheetApp.getActiveSpreadsheet(), now = new Date();
  function sheet(name, headers) {
    var s = ss.getSheetByName(name) || ss.insertSheet(name);
    if (s.getLastRow() === 0) { s.appendRow(headers); s.setFrozenRows(1); }
    return s;
  }
  if (d.type === 'event') {
    sheet('Events', ['Received','Username','Attempt','Round','Event','Time']).appendRow([now, d.username, d.attempt, d.round, d.event, d.at]);
    return ContentService.createTextOutput('ok');
  }
  var tag = [d.username, d.attempt, d.round];
  sheet('Rounds', ['Received','Ended by','Username','Attempt','Round','Total attempts so far','Restarted from (map / end page)','Continued after end page (round > 1)','Game started','Round started','Round ended','Score',
      'Badges (count)','All badges collected','New badges this attempt','Badges by attempt','Station route (completed)','Stations opened (order)','Concept route (main)','Concept route (CoMS)','Station times (s)','CoMS progress','Replays (new attempts)',
      'Hints used','Questions answered','Stations (total)','Stations where a hint was used','Station with most hints'])
    .appendRow([now, d.reason, d.username, d.attempt, d.round, d.totalAttemptsSoFar, d.restartedFrom, d.continuedAfterEndPage ? 'yes' : 'no', d.startedAt, d.roundStartedAt, d.endedAt, d.score,
      d.badges.length, d.badges.join('; '), d.newBadgesThisAttempt.join('; '), d.badgesByAttempt, d.route.join(' > '), d.stationsOpenedOrder.join(' > '), d.conceptRoute.join(' > '), d.comsRoute.join(' > '),
      JSON.stringify(d.stationTimes), d.comsProgress.join('; '), d.replays, d.hintsUsed, d.questionsAnswered, d.stationsTotal, d.stationsWithHints, d.topHintStation]);
  var qs = sheet('Questions', ['Received','Username','Attempt','Round','Rank (1 = longest)','Station','Concept','Type','Question','Time (s)','Hint used','Correct']);
  d.questions.forEach(function(q){ qs.appendRow([now].concat(tag, [q.rank, q.station, q.concept, q.kind, q.q, q.seconds, q.hint ? 'yes':'no', q.correct ? 'yes':'no'])); });
  var cs = sheet('Concept visits', ['Received','Username','Attempt','Round','Station','Concept','Type','Flashcards time (s)','Questions time (s)','Total time (s)','Wrong answers','Hint used','Completed']);
  d.concepts.forEach(function(c){ cs.appendRow([now].concat(tag, [c.station, c.concept, c.kind, c.flashSeconds, c.questionSeconds, c.totalSeconds, c.wrong, c.hint ? 'yes':'no', c.completed ? 'yes':'no'])); });
  var st = sheet('Stations', ['Received','Username','Attempt','Round (cumulative to here)','Station','Type','Time (s)','Questions attempted','Correct','Wrong','Minimum questions needed to pass','Hints used','Concepts completed','Concepts total']);
  d.stations.forEach(function(r){ st.appendRow([now].concat(tag, [r.station, r.kind, r.seconds, r.attempted, r.correct, r.wrong, r.minRequired, r.hints, r.conceptsCompleted, r.conceptsTotal])); });
  var cc = sheet('Concept summary', ['Received','Username','Attempt','Round (cumulative to here)','Station','Concept','Type','Total time (s)','Flashcards time (s)','Questions time (s)','Questions attempted','Correct','Wrong','Hints used','Visits']);
  d.conceptSummary.forEach(function(r){ cc.appendRow([now].concat(tag, [r.station, r.concept, r.kind, r.totalSeconds, r.flashSeconds, r.questionSeconds, r.attempted, r.correct, r.wrong, r.hints, r.visits])); });
  return ContentService.createTextOutput('ok');
}
