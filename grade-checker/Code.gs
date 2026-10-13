// Deploy this separate grade-checker project as Me with access Anyone.
// Keep the spreadsheet and rosters private.
const TRACKER_ID = '1IwfPam3I-o4cFu09R9Z6U86RK59LF96ZZjlEqArPAoM';
const INCLUDED_LESSONS = [["01","Programming for a Purpose"],["02","Plotting Shapes"],["03","Drawing in Game Lab"],["04","Shapes and Parameters"],["05","Variables"],["06","Random Numbers"],["08","Sprites"],["09","Sprite Properties"],["10","Text"],["12","The Draw Loop"],["13","Sprite Movement"],["15","Conditionals"],["16","Keyboard Input"],["17","Mouse Input"],["19","Velocity"],["20","Collision Detection"],["22","Complex Sprite Movement"],["23","Collisions"],["25","Functions"]];

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Scores')
    .setTitle('IST Game Design Practice-Lab Scores');
}

function lookupDenied() {
  throw new Error('Student ID and last name could not be verified. Check both entries or ask your teacher.');
}

function normalizedName(value) {
  return String(value || '').trim().replace(/\s+/g, ' ').toLowerCase();
}

function getMyScores(studentId, lastName) {
  const id = String(studentId || '').trim();
  const last = normalizedName(lastName);
  if (!/^[0-9]{1,30}$/.test(id) || !last || last.length > 100) lookupDenied();
  // Limit repeated lookup attempts without saving a public roster or credentials.
  const cache = CacheService.getScriptCache();
  const userKey = Session.getTemporaryActiveUserKey() || 'shared-anonymous';
  const bucket = 'gradeLookup:' + userKey;
  const limit = userKey === 'shared-anonymous' ? 180 : 15;
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const attempts = Number(cache.get(bucket) || 0);
    if (attempts >= limit) throw new Error('Too many lookup attempts. Please wait a minute and try again.');
    cache.put(bucket, String(attempts + 1), 60);
  } finally { lock.releaseLock(); }
  const book = SpreadsheetApp.openById(TRACKER_ID);
  const matches = [];
  [['1st Block Roster','1st Block'],['4th Block Roster','4th Block']].forEach(function(entry) {
    const roster = book.getSheetByName(entry[0]);
    if (!roster || roster.getLastRow() < 2) return;
    roster.getRange(2,1,roster.getLastRow()-1,4).getDisplayValues().forEach(function(row) {
      if (String(row[0]).trim() === id && String(row[3]).trim() === entry[1])
        matches.push({first:row[2],last:row[1],period:entry[1]});
    });
  });
  if (matches.length !== 1 || normalizedName(matches[0].last) !== last) lookupDenied();
  const student = matches[0];
  const validNames = [
    normalizedName(student.first + ' ' + student.last),
    normalizedName(student.last + ', ' + student.first)
  ];
  const sheet = book.getSheetByName('Form Responses 1');
  if (!sheet) throw new Error('Scores are temporarily unavailable. Please try again later.');
  const header = sheet.getRange(1,1,1,5).getDisplayValues()[0];
  if (header.join('|') !== 'Timestamp|Student Name|Class Period|Practice Lab|Score')
    throw new Error('Scores are temporarily unavailable. Please ask your teacher.');
  const highest = {};
  if (sheet.getLastRow() > 1) {
    sheet.getRange(2,1,sheet.getLastRow()-1,5).getValues().forEach(function(row) {
      if (String(row[2]).trim() !== student.period || !validNames.includes(normalizedName(row[1]))) return;
      const lesson = String(row[3]).match(/^Lesson ([0-9]{2})(:| [|])/);
      const score = Number(row[4]);
      if (!lesson || row[4] === '' || !Number.isFinite(score) || score < 0 || score > 100) return;
      const n = lesson[1];
      if (!INCLUDED_LESSONS.some(function(item) { return item[0] === n; })) return;
      if (highest[n] === undefined || score > highest[n]) highest[n] = score;
    });
  }
  const labs = INCLUDED_LESSONS.map(function(item,i) {
    return {lesson:item[0],title:item[1],half:i<9?1:2,
      completed:highest[item[0]] !== undefined,
      score:highest[item[0]] === undefined?null:highest[item[0]]};
  });
  function average(list) {
    return Math.round(list.reduce(function(sum,lab) { return sum + (lab.completed?lab.score:0); },0) / list.length * 10) / 10;
  }
  // Return this verified student's grades only, never the rosters/access map.
  return {name:student.first+' '+student.last,period:student.period,labs:labs,
    completed:labs.filter(function(lab) { return lab.completed; }).length,
    firstAverage:average(labs.slice(0,9)),secondAverage:average(labs.slice(9))};
}
