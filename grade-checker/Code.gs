// Deploy this separate grade-checker project using the school account.
// Execute as: Me. Access: users in benhillschools.org only.
// Never deploy as Anyone. Keep Grade Lookup Access inside the private tracker.
const TRACKER_ID = '1IwfPam3I-o4cFu09R9Z6U86RK59LF96ZZjlEqArPAoM';
const SCHOOL_DOMAIN = 'benhillschools.org';
const INCLUDED_LESSONS = [["01","Programming for a Purpose"],["02","Plotting Shapes"],["03","Drawing in Game Lab"],["04","Shapes and Parameters"],["05","Variables"],["06","Random Numbers"],["08","Sprites"],["09","Sprite Properties"],["10","Text"],["12","The Draw Loop"],["13","Sprite Movement"],["15","Conditionals"],["16","Keyboard Input"],["17","Mouse Input"],["19","Velocity"],["20","Collision Detection"],["22","Complex Sprite Movement"],["23","Collisions"],["25","Functions"]];

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Scores')
    .setTitle('IST Game Design Practice-Lab Scores');
}

// Run once as the teacher. Enter Student ID and the student's school email
// in this PRIVATE tab. IDs and emails are never sent to GitHub or the browser.
function setupLookupAccess() {
  const book = SpreadsheetApp.openById(TRACKER_ID);
  let sheet = book.getSheetByName('Grade Lookup Access');
  if (!sheet) {
    sheet = book.insertSheet('Grade Lookup Access');
    sheet.getRange(1,1,1,2).setValues([['Student ID','School Email']]);
    sheet.getRange('A:A').setNumberFormat('@');
    sheet.setFrozenRows(1);
    sheet.getRange(1,1,1,2).setBackground('#48216d').setFontColor('#ffffff').setFontWeight('bold');
    sheet.setColumnWidth(1,150);
    sheet.setColumnWidth(2,300);
  }
}

function lookupDenied() {
  throw new Error('We could not verify access. Sign in with your school account and check your Student ID. Ask your teacher if this continues.');
}

function normalizedName(value) {
  return String(value || '').trim().replace(/\s+/g, ' ').toLowerCase();
}

function getMyScores(studentId) {
  const email = String(Session.getActiveUser().getEmail() || '').trim().toLowerCase();
  const id = String(studentId || '').trim();
  if (!email || !email.endsWith('@' + SCHOOL_DOMAIN) || !/^[a-zA-Z0-9-]{1,30}$/.test(id)) lookupDenied();
  const book = SpreadsheetApp.openById(TRACKER_ID);
  const access = book.getSheetByName('Grade Lookup Access');
  if (!access || access.getLastRow() < 2) lookupDenied();
  const links = access.getRange(2,1,access.getLastRow()-1,2).getDisplayValues();
  const idLinks = links.filter(function(row) { return String(row[0]).trim() === id; });
  const emailLinks = links.filter(function(row) { return String(row[1]).trim().toLowerCase() === email; });
  if (idLinks.length !== 1 || emailLinks.length !== 1 || String(idLinks[0][1]).trim().toLowerCase() !== email) lookupDenied();

  const matches = [];
  [['1st Block Roster','1st Block'],['4th Block Roster','4th Block']].forEach(function(entry) {
    const roster = book.getSheetByName(entry[0]);
    if (!roster || roster.getLastRow() < 2) return;
    roster.getRange(2,1,roster.getLastRow()-1,4).getDisplayValues().forEach(function(row) {
      if (String(row[0]).trim() === id && String(row[3]).trim() === entry[1])
        matches.push({first:row[2],last:row[1],period:entry[1]});
    });
  });
  if (matches.length !== 1) lookupDenied();
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
