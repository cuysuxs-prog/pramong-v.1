// ============================================================
// Pramong Taidin - Code.gs
// Google Apps Script สำหรับเชื่อม Google Sheets
// ============================================================
// ⚙️ แก้ไขค่าเหล่านี้ก่อนใช้งาน
var SPREADSHEET_ID   = '1Jfsb-Pix5z1m-jdft9fYAI9dKseAxv5ICFx1cf0YYNA'; // ID ของ Google Sheets
var TEACHER_PASSWORD = '1234';                       // รหัสผ่านอาจารย์

// ชื่อชีต
var SHEET_PLAYERS   = 'Players';
var SHEET_QUESTIONS = 'Questions';
var SHEET_ANSWERS   = 'Answers';
var SHEET_SESSIONS  = 'GameSessions';
var SHEET_AQUATYPE  = 'AquaTypeResults';

// ============================================================
// Entry Points
// ============================================================
function doGet(e) {
  try {
    var action = e.parameter.action || '';
    var result = handleGet(action, e.parameter);
    return respond(result);
  } catch (err) {
    return respond({ success: false, error: err.message });
  }
}

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var result = handlePost(data.action || '', data);
    return respond(result);
  } catch (err) {
    return respond({ success: false, error: err.message });
  }
}

function respond(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// GET Actions (อ่านข้อมูล)
// ============================================================
function handleGet(action, p) {
  switch (action) {
    case 'getGameState':      return getGameState(p.gameId);
    case 'getCurrentQuestion':return getCurrentQuestion(p.gameId);
    case 'getLeaderboard':    return getLeaderboard(p.gameId, parseInt(p.limit) || 10);
    case 'getMyScore':        return getMyScore(p.gameId, p.playerId);
    case 'getPlayerCount':    return getPlayerCount(p.gameId);
    case 'getQuestionSets':   return getQuestionSets(p.password);
    case 'getQuestionSet':    return getQuestionSetById(p.setId, p.password);
    case 'getGameResults':    return getGameResults(p.gameId, p.password);
    case 'getAnswerStats':    return getAnswerStats(p.gameId, parseInt(p.qNum), p.password);
    case 'getTextAnswers':    return getTextAnswers(p.gameId, parseInt(p.qNum));
    case 'getAquaTypeStats':  return getAquaTypeStats();
    case 'getGameByPin':      return getGameByPin(p.pin);
    default: return { success: false, error: 'unknown action: ' + action };
  }
}

// ============================================================
// POST Actions (เขียนข้อมูล)
// ============================================================
function handlePost(action, data) {
  switch (action) {
    // นักศึกษา
    case 'joinGame':      return joinGame(data);
    case 'submitAnswer':  return submitAnswer(data);

    // อาจารย์
    case 'createGame':       return createGame(data);
    case 'nextQuestion':     return nextQuestion(data);
    case 'prevQuestion':     return prevQuestion(data);
    case 'showAnswer':       return showAnswer(data);
    case 'showLeaderboard':  return showLeaderboardCmd(data);
    case 'closeAnswers':     return setStatus(data, 'closed');
    case 'openAnswers':      return setStatus(data, 'playing');
    case 'endGame':          return endGame(data);
    case 'resetGame':        return resetGame(data);
    case 'saveQuestionSet':  return saveQuestionSet(data);
    case 'updateQuestionSet':return updateQuestionSet(data);
    case 'deleteQuestionSet':return deleteQuestionSet(data);
    case 'copyQuestionSet':  return copyQuestionSet(data);
    case 'submitAquaType':   return submitAquaType(data);
    default: return { success: false, error: 'unknown action: ' + action };
  }
}

// ============================================================
// Helper Functions
// ============================================================
function getSheet(name) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sh = ss.getSheetByName(name);
  if (!sh) throw new Error('ไม่พบชีต: ' + name);
  return sh;
}

function auth(password) {
  if (password !== TEACHER_PASSWORD) throw new Error('รหัสผ่านไม่ถูกต้อง');
}

function uid(prefix) {
  return prefix + '_' + Date.now() + '_' + Math.random().toString(36).substr(2, 6).toUpperCase();
}

function genPIN() {
  return Math.floor(100000 + Math.random() * 900000).toString();
}

// ============================================================
// GameSessions Column Layout (0-indexed)
// [0]GameID [1]GamePIN [2]QuestionSetID [3]QuestionSetName
// [4]StartTime [5]EndTime [6]TotalPlayers [7]TotalQuestions
// [8]CurrentQuestionNumber [9]Status
// ============================================================
function getSessionRow(gameId) {
  var sh = getSheet(SHEET_SESSIONS);
  var data = sh.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === gameId) return { sheet: sh, row: i + 1, data: data[i] };
  }
  throw new Error('ไม่พบเกม: ' + gameId);
}

function getSessionByPIN(pin) {
  var sh = getSheet(SHEET_SESSIONS);
  var data = sh.getDataRange().getValues();
  for (var i = data.length - 1; i >= 1; i--) {
    if (data[i][1].toString() === pin.toString() && data[i][9] !== 'ended') {
      return { sheet: sh, row: i + 1, data: data[i] };
    }
  }
  return null;
}

function rowToSession(d) {
  return {
    gameId:                d[0],
    pin:                   d[1],
    questionSetId:         d[2],
    questionSetName:       d[3],
    startTime:             d[4],
    endTime:               d[5],
    totalPlayers:          parseInt(d[6]) || 0,
    totalQuestions:        parseInt(d[7]) || 0,
    currentQuestionNumber: parseInt(d[8]) || 0,
    status:                d[9]
  };
}

// ============================================================
// Student: joinGame
// ============================================================
function joinGame(data) {
  var pin    = data.pin;
  var player = data.player;

  var sess = getSessionByPIN(pin);
  if (!sess) return { success: false, error: 'ไม่พบห้องเกม หรือเกมสิ้นสุดแล้ว (PIN: ' + pin + ')' };

  var session = rowToSession(sess.data);

  // ตรวจว่าเคยสมัครด้วย studentId เดิมหรือยัง
  var plSh   = getSheet(SHEET_PLAYERS);
  var plData = plSh.getDataRange().getValues();
  for (var j = 1; j < plData.length; j++) {
    if (plData[j][1] === session.gameId && plData[j][5] === player.studentId) {
      return {
        success:  true,
        gameId:   session.gameId,
        playerId: plData[j][2],
        nickname: plData[j][4],
        message:  'ยินดีต้อนรับกลับ!'
      };
    }
  }

  // สร้าง PlayerID ใหม่
  var playerId = uid('P');
  plSh.appendRow([
    new Date(),          // [0] Timestamp
    session.gameId,      // [1] GameID
    playerId,            // [2] PlayerID
    player.fullName,     // [3] FullName
    player.nickname,     // [4] Nickname
    player.studentId,    // [5] StudentID
    player.program,      // [6] Program
    player.year,         // [7] Year
    player.province,     // [8] Province
    player.expectation,  // [9] Expectation
    0                    // [10] TotalScore
  ]);

  // อัพเดทจำนวนผู้เล่น
  var newCount = (session.totalPlayers || 0) + 1;
  sess.sheet.getRange(sess.row, 7).setValue(newCount);

  return {
    success:  true,
    gameId:   session.gameId,
    playerId: playerId,
    nickname: player.nickname,
    message:  'เข้าร่วมเกมสำเร็จ! ยินดีต้อนรับ ' + player.nickname
  };
}

// ============================================================
// Student: getCurrentQuestion
// ============================================================
function getCurrentQuestion(gameId) {
  var sess = getSessionRow(gameId);
  var s    = rowToSession(sess.data);

  var qData = null;
  if (s.currentQuestionNumber > 0) {
    var qSh   = getSheet(SHEET_QUESTIONS);
    var qRows = qSh.getDataRange().getValues();
    for (var i = 1; i < qRows.length; i++) {
      if (qRows[i][0] === s.questionSetId && parseInt(qRows[i][2]) === s.currentQuestionNumber) {
        var showAnswer = (s.status === 'showing_answer' || s.status === 'leaderboard');
        qData = {
          setId:         qRows[i][0],
          setName:       qRows[i][3],
          number:        parseInt(qRows[i][2]),
          type:          qRows[i][3] || 'multiple_choice', // fix below
          question:      qRows[i][4],
          imageUrl:      qRows[i][5],
          optionA:       qRows[i][6],
          optionB:       qRows[i][7],
          optionC:       qRows[i][8],
          optionD:       qRows[i][9],
          correctAnswer: showAnswer ? qRows[i][10] : null,
          explanation:   showAnswer ? qRows[i][11] : null,
          timeLimit:     parseInt(qRows[i][12]) || 30,
          points:        parseInt(qRows[i][13]) || 1000
        };
        // type column is index 3 in Questions sheet
        qData.type    = qRows[i][3];
        qData.setName = qRows[i][1];
        break;
      }
    }
  }

  return {
    success:  true,
    session:  s,
    question: qData
  };
}

// ============================================================
// Student: submitAnswer
// ============================================================
function submitAnswer(data) {
  var gameId         = data.gameId;
  var playerId       = data.playerId;
  var questionNumber = parseInt(data.questionNumber);
  var answer         = data.answer;
  var responseTime   = parseFloat(data.responseTime) || 0; // milliseconds

  // ตรวจซ้ำ
  var ansSh   = getSheet(SHEET_ANSWERS);
  var ansData = ansSh.getDataRange().getValues();
  for (var i = 1; i < ansData.length; i++) {
    if (ansData[i][1] === gameId && ansData[i][2] === playerId &&
        parseInt(ansData[i][6]) === questionNumber) {
      return { success: false, alreadyAnswered: true, error: 'ส่งคำตอบข้อนี้แล้ว' };
    }
  }

  // ดึง session
  var sess = getSessionRow(gameId);
  var s    = rowToSession(sess.data);
  if (s.status !== 'playing') {
    return { success: false, error: 'ปิดรับคำตอบแล้ว (status: ' + s.status + ')' };
  }

  // ดึงคำถาม
  var qSh   = getSheet(SHEET_QUESTIONS);
  var qData = qSh.getDataRange().getValues();
  var q     = null;
  for (var j = 1; j < qData.length; j++) {
    if (qData[j][0] === s.questionSetId && parseInt(qData[j][2]) === questionNumber) {
      q = qData[j];
      break;
    }
  }
  if (!q) return { success: false, error: 'ไม่พบคำถาม' };

  var qType         = q[3];
  var correctAnswer = q[10];
  var timeLimit     = parseInt(q[12]) || 30;
  var maxPoints     = parseInt(q[13]) || 1000;

  // คำนวณคะแนน
  var isCorrect = false;
  var score     = 0;
  if (qType === 'multiple_choice' || qType === 'true_false') {
    isCorrect = answer.toString().toUpperCase() === correctAnswer.toString().toUpperCase();
    if (isCorrect) {
      var timeRatio = Math.max(0, 1 - (responseTime / (timeLimit * 1000)));
      score = Math.round(maxPoints * (0.5 + 0.5 * timeRatio));
    }
  } else {
    // Poll / Brainstorm / Word Cloud — ให้คะแนนมีส่วนร่วม
    isCorrect = true;
    score     = 100;
  }

  // ดึงข้อมูลผู้เล่น
  var plSh   = getSheet(SHEET_PLAYERS);
  var plData = plSh.getDataRange().getValues();
  var nickname  = '';
  var studentId = '';
  var plRow     = -1;
  for (var k = 1; k < plData.length; k++) {
    if (plData[k][1] === gameId && plData[k][2] === playerId) {
      nickname  = plData[k][4];
      studentId = plData[k][5];
      plRow     = k + 1;
      break;
    }
  }

  // บันทึกคำตอบ
  ansSh.appendRow([
    new Date(),      // [0]  Timestamp
    gameId,          // [1]  GameID
    playerId,        // [2]  PlayerID
    nickname,        // [3]  Nickname
    studentId,       // [4]  StudentID
    q[1],            // [5]  QuestionSetName
    questionNumber,  // [6]  QuestionNumber
    q[4],            // [7]  Question text
    answer,          // [8]  Answer
    correctAnswer,   // [9]  CorrectAnswer
    isCorrect,       // [10] IsCorrect
    responseTime,    // [11] ResponseTime (ms)
    score            // [12] Score
  ]);

  // อัพเดทคะแนนรวม
  if (plRow > 0) {
    var curScore = parseInt(plData[plRow - 1][10]) || 0;
    plSh.getRange(plRow, 11).setValue(curScore + score);
  }

  return {
    success:       true,
    isCorrect:     isCorrect,
    score:         score,
    correctAnswer: correctAnswer,
    message:       isCorrect ? 'ตอบถูก! +' + score + ' คะแนน' : 'ตอบผิด!'
  };
}

// ============================================================
// Student: getLeaderboard
// ============================================================
function getLeaderboard(gameId, limit) {
  limit = limit || 10;
  var plSh   = getSheet(SHEET_PLAYERS);
  var plData = plSh.getDataRange().getValues();
  var list   = [];
  for (var i = 1; i < plData.length; i++) {
    if (plData[i][1] === gameId) {
      list.push({
        nickname:   plData[i][4],
        totalScore: parseInt(plData[i][10]) || 0
      });
    }
  }
  list.sort(function (a, b) { return b.totalScore - a.totalScore; });
  list.forEach(function (p, idx) { p.rank = idx + 1; });
  return { success: true, leaderboard: list.slice(0, limit), totalPlayers: list.length };
}

// ============================================================
// Student: getMyScore
// ============================================================
function getMyScore(gameId, playerId) {
  var plSh   = getSheet(SHEET_PLAYERS);
  var plData = plSh.getDataRange().getValues();
  var me     = null;
  for (var i = 1; i < plData.length; i++) {
    if (plData[i][1] === gameId && plData[i][2] === playerId) {
      me = { nickname: plData[i][4], totalScore: parseInt(plData[i][10]) || 0 };
      break;
    }
  }
  if (!me) return { success: false, error: 'ไม่พบผู้เล่น' };

  var scores = [];
  for (var j = 1; j < plData.length; j++) {
    if (plData[j][1] === gameId) scores.push(parseInt(plData[j][10]) || 0);
  }
  scores.sort(function (a, b) { return b - a; });
  var rank = scores.indexOf(me.totalScore) + 1;

  return { success: true, nickname: me.nickname, score: me.totalScore, rank: rank, totalPlayers: scores.length };
}

// ============================================================
// Display: getGameState / getPlayerCount
// ============================================================
function getGameByPin(pin) {
  var sess = getSessionByPIN(pin);
  if (!sess) return { success: false, error: 'ไม่พบเกมที่มี PIN นี้' };
  return { success: true, gameId: sess.data[0] };
}

function getGameState(gameId) {
  var sess = getSessionRow(gameId);
  return { success: true, session: rowToSession(sess.data) };
}

function getPlayerCount(gameId) {
  var plSh   = getSheet(SHEET_PLAYERS);
  var plData = plSh.getDataRange().getValues();
  var count  = 0;
  var nicks  = [];
  for (var i = 1; i < plData.length; i++) {
    if (plData[i][1] === gameId) { count++; nicks.push(plData[i][4]); }
  }
  return { success: true, count: count, players: nicks };
}

// ============================================================
// Display: getTextAnswers (Brainstorm / Word Cloud live feed)
// ============================================================
function getTextAnswers(gameId, questionNumber) {
  var ansSh   = getSheet(SHEET_ANSWERS);
  var ansData = ansSh.getDataRange().getValues();
  var answers = [];
  for (var i = 1; i < ansData.length; i++) {
    if (ansData[i][1] === gameId && parseInt(ansData[i][6]) === questionNumber) {
      answers.push({
        nickname: ansData[i][3],   // ชื่อเล่นเท่านั้น ไม่ส่ง studentId / fullName
        answer:   String(ansData[i][8])
      });
    }
  }
  return { success: true, answers: answers };
}

// ============================================================
// Teacher: getAnswerStats (สถิติคำตอบรายข้อ)
// ============================================================
function getAnswerStats(gameId, questionNumber, password) {
  auth(password);
  var ansSh   = getSheet(SHEET_ANSWERS);
  var ansData = ansSh.getDataRange().getValues();
  var stats   = { A: 0, B: 0, C: 0, D: 0, text: [], total: 0 };
  for (var i = 1; i < ansData.length; i++) {
    if (ansData[i][1] === gameId && parseInt(ansData[i][6]) === questionNumber) {
      var ans = ansData[i][8].toString().toUpperCase();
      stats.total++;
      if (ans === 'A') stats.A++;
      else if (ans === 'B') stats.B++;
      else if (ans === 'C') stats.C++;
      else if (ans === 'D') stats.D++;
      else stats.text.push(ansData[i][8]);
    }
  }
  return { success: true, stats: stats };
}

// ============================================================
// Teacher: createGame
// ============================================================
function createGame(data) {
  auth(data.password);
  var setId = data.setId;

  var qSh   = getSheet(SHEET_QUESTIONS);
  var qData = qSh.getDataRange().getValues();
  var count = 0;
  var setName = '';
  for (var i = 1; i < qData.length; i++) {
    if (qData[i][0] === setId) { count++; setName = qData[i][1]; }
  }
  if (count === 0) return { success: false, error: 'ไม่พบคำถามในชุดนี้' };

  var gameId = uid('G');
  var pin    = genPIN();

  getSheet(SHEET_SESSIONS).appendRow([
    gameId,    // [0] GameID
    pin,       // [1] GamePIN
    setId,     // [2] QuestionSetID
    setName,   // [3] QuestionSetName
    new Date(),// [4] StartTime
    '',        // [5] EndTime
    0,         // [6] TotalPlayers
    count,     // [7] TotalQuestions
    0,         // [8] CurrentQuestionNumber
    'waiting'  // [9] Status
  ]);

  return { success: true, gameId: gameId, pin: pin, totalQuestions: count, questionSetName: setName };
}

// ============================================================
// Teacher: game flow commands
// ============================================================
function nextQuestion(data) {
  auth(data.password);
  var sess = getSessionRow(data.gameId);
  var s    = rowToSession(sess.data);
  if (s.currentQuestionNumber >= s.totalQuestions) return { success: false, error: 'ถึงคำถามข้อสุดท้ายแล้ว' };
  var newNum = s.currentQuestionNumber + 1;
  sess.sheet.getRange(sess.row, 9).setValue(newNum);
  sess.sheet.getRange(sess.row, 10).setValue('playing');
  return { success: true, currentQuestion: newNum, status: 'playing' };
}

function prevQuestion(data) {
  auth(data.password);
  var sess = getSessionRow(data.gameId);
  var s    = rowToSession(sess.data);
  if (s.currentQuestionNumber <= 1) return { success: false, error: 'อยู่ที่ข้อแรกแล้ว' };
  var newNum = s.currentQuestionNumber - 1;
  sess.sheet.getRange(sess.row, 9).setValue(newNum);
  sess.sheet.getRange(sess.row, 10).setValue('playing');
  return { success: true, currentQuestion: newNum, status: 'playing' };
}

function showAnswer(data) {
  auth(data.password);
  return setStatus(data, 'showing_answer');
}

function showLeaderboardCmd(data) {
  auth(data.password);
  return setStatus(data, 'leaderboard');
}

function setStatus(data, newStatus) {
  if (data.password) auth(data.password);
  var sess = getSessionRow(data.gameId);
  sess.sheet.getRange(sess.row, 10).setValue(newStatus);
  return { success: true, status: newStatus };
}

function endGame(data) {
  auth(data.password);
  var sess = getSessionRow(data.gameId);
  sess.sheet.getRange(sess.row, 6).setValue(new Date()); // EndTime
  sess.sheet.getRange(sess.row, 10).setValue('ended');
  return { success: true, status: 'ended' };
}

function resetGame(data) {
  auth(data.password);
  var gameId = data.gameId;

  // ลบผู้เล่น
  deleteRowsByGameId(SHEET_PLAYERS, gameId, 1);
  // ลบคำตอบ
  deleteRowsByGameId(SHEET_ANSWERS, gameId, 1);

  // รีเซ็ต session
  var sess = getSessionRow(gameId);
  sess.sheet.getRange(sess.row, 5).setValue('');  // EndTime
  sess.sheet.getRange(sess.row, 7).setValue(0);   // TotalPlayers
  sess.sheet.getRange(sess.row, 9).setValue(0);   // CurrentQuestionNumber
  sess.sheet.getRange(sess.row, 10).setValue('waiting'); // Status

  return { success: true, message: 'รีเซ็ตเกมสำเร็จ' };
}

function deleteRowsByGameId(sheetName, gameId, gameIdCol) {
  var sh   = getSheet(sheetName);
  var data = sh.getDataRange().getValues();
  for (var i = data.length - 1; i >= 1; i--) {
    if (data[i][gameIdCol] === gameId) sh.deleteRow(i + 1);
  }
}

// ============================================================
// Teacher: Question Bank
// ============================================================
function getQuestionSets(password) {
  auth(password);
  var qSh   = getSheet(SHEET_QUESTIONS);
  var qData = qSh.getDataRange().getValues();
  var map   = {};
  for (var i = 1; i < qData.length; i++) {
    var sid = qData[i][0];
    if (!sid) continue;
    if (!map[sid]) map[sid] = { id: sid, name: qData[i][1], count: 0 };
    map[sid].count++;
  }
  return { success: true, sets: Object.values(map) };
}

function getQuestionSetById(setId, password) {
  auth(password);
  var qSh   = getSheet(SHEET_QUESTIONS);
  var qData = qSh.getDataRange().getValues();
  var qs    = [];
  for (var i = 1; i < qData.length; i++) {
    if (qData[i][0] === setId) {
      qs.push({
        setId:         qData[i][0],
        setName:       qData[i][1],
        number:        parseInt(qData[i][2]),
        type:          qData[i][3],
        question:      qData[i][4],
        imageUrl:      qData[i][5],
        optionA:       qData[i][6],
        optionB:       qData[i][7],
        optionC:       qData[i][8],
        optionD:       qData[i][9],
        correctAnswer: qData[i][10],
        explanation:   qData[i][11],
        timeLimit:     parseInt(qData[i][12]) || 30,
        points:        parseInt(qData[i][13]) || 1000
      });
    }
  }
  qs.sort(function (a, b) { return a.number - b.number; });
  return { success: true, questions: qs };
}

function saveQuestionSet(data) {
  auth(data.password);
  var setId   = data.setId || uid('QS');
  var setName = data.setName;
  var qSh     = getSheet(SHEET_QUESTIONS);
  data.questions.forEach(function (q, idx) {
    qSh.appendRow([
      setId, setName, idx + 1,
      q.type || 'multiple_choice',
      q.question, q.imageUrl || '',
      q.optionA || '', q.optionB || '', q.optionC || '', q.optionD || '',
      q.correctAnswer || '', q.explanation || '',
      q.timeLimit || 30, q.points || 1000
    ]);
  });
  return { success: true, setId: setId, message: 'บันทึกชุดคำถามสำเร็จ' };
}

function updateQuestionSet(data) {
  auth(data.password);
  deleteRowsBySetId(data.setId);
  var qSh = getSheet(SHEET_QUESTIONS);
  data.questions.forEach(function (q, idx) {
    qSh.appendRow([
      data.setId, data.setName, idx + 1,
      q.type || 'multiple_choice',
      q.question, q.imageUrl || '',
      q.optionA || '', q.optionB || '', q.optionC || '', q.optionD || '',
      q.correctAnswer || '', q.explanation || '',
      q.timeLimit || 30, q.points || 1000
    ]);
  });
  return { success: true, message: 'อัพเดทชุดคำถามสำเร็จ' };
}

function deleteQuestionSet(data) {
  auth(data.password);
  deleteRowsBySetId(data.setId);
  return { success: true, message: 'ลบชุดคำถามสำเร็จ' };
}

function copyQuestionSet(data) {
  auth(data.password);
  var newId   = uid('QS');
  var newName = data.newName || (data.setName + ' (สำเนา)');
  var qSh     = getSheet(SHEET_QUESTIONS);
  var qData   = qSh.getDataRange().getValues();
  for (var i = 1; i < qData.length; i++) {
    if (qData[i][0] === data.setId) {
      var r = qData[i].slice();
      r[0] = newId;
      r[1] = newName;
      qSh.appendRow(r);
    }
  }
  return { success: true, newSetId: newId, newSetName: newName };
}

function deleteRowsBySetId(setId) {
  var qSh   = getSheet(SHEET_QUESTIONS);
  var qData = qSh.getDataRange().getValues();
  for (var i = qData.length - 1; i >= 1; i--) {
    if (qData[i][0] === setId) qSh.deleteRow(i + 1);
  }
}

// ============================================================
// Teacher: getGameResults (Export)
// ============================================================
function getGameResults(gameId, password) {
  auth(password);
  var ansSh   = getSheet(SHEET_ANSWERS);
  var ansData = ansSh.getDataRange().getValues();
  var rows    = [];
  for (var i = 1; i < ansData.length; i++) {
    if (ansData[i][1] === gameId) {
      rows.push({
        timestamp:      ansData[i][0],
        nickname:       ansData[i][3],
        studentId:      ansData[i][4],
        questionNumber: ansData[i][6],
        question:       ansData[i][7],
        answer:         ansData[i][8],
        correctAnswer:  ansData[i][9],
        isCorrect:      ansData[i][10],
        responseTime:   ansData[i][11],
        score:          ansData[i][12]
      });
    }
  }
  return { success: true, results: rows };
}

// ============================================================
// ตัวอย่างข้อมูลชุดคำถามเริ่มต้น (รันครั้งเดียวเพื่อเติมข้อมูล)
// ============================================================
function insertSampleQuestions() {
  var qSh    = getSheet(SHEET_QUESTIONS);
  var setId  = 'QS_SAMPLE_001';
  var name   = 'ความรู้ประมงเบื้องต้น';
  var sample = [
    ['multiple_choice','ปลาค้อ (Siamese algae eater) จัดอยู่ในวงศ์ใด?','','Cyprinidae','Clariidae','Channidae','Siluridae','A','ปลาค้อจัดอยู่ในวงศ์ Cyprinidae หรือวงศ์ปลาตะเพียน',30,1000],
    ['multiple_choice','การเพาะเลี้ยงสัตว์น้ำในระบบปิด (RAS) ย่อมาจากอะไร?','','Recirculating Aquaculture System','Rapid Aquatic System','Reverse Aquaculture Setup','Renewable Aqua Source','A','RAS = Recirculating Aquaculture System ระบบหมุนเวียนน้ำ',30,1000],
    ['true_false','ค่า pH ที่เหมาะสมสำหรับปลาน้ำจืดส่วนใหญ่อยู่ในช่วง 6.5–8.5','','ถูก','ผิด','','','A','ปลาน้ำจืดส่วนใหญ่ชอบ pH 6.5–8.5',20,1000],
    ['multiple_choice','ออกซิเจนละลายน้ำ (DO) ที่เหมาะสมสำหรับการเพาะเลี้ยงปลาคือเท่าใด?','','>5 mg/L','1-2 mg/L','10-15 mg/L','0.5 mg/L','A','DO ควรสูงกว่า 5 mg/L เพื่อสุขภาพที่ดีของปลา',30,1000],
    ['multiple_choice','Smart Aquaculture หมายถึงอะไร?','','การเพาะเลี้ยงสัตว์น้ำด้วยเทคโนโลยีอัจฉริยะ','การเพาะเลี้ยงปลาฉลาม','ระบบประมงน้ำลึก','การทำฟาร์มบนบก','A','Smart Aquaculture คือการนำ IoT และ AI มาใช้ในการเพาะเลี้ยงสัตว์น้ำ',30,1000],
    ['poll','คุณสนใจเรียนด้านประมงสาขาใดมากที่สุด?','','การเพาะเลี้ยงสัตว์น้ำ','การประมงทะเล','แปรรูปอาหารทะเล','เทคโนโลยีประมง','','',30,0],
    ['brainstorm','บอกสิ่งที่คาดหวังจากการเรียนประมงมา 1 อย่าง','','','','','','','',60,100],
    ['word_cloud','คิดถึงคำว่า "ประมง" แล้วนึกถึงคำใดบ้าง? พิมพ์มา 1-2 คำ','','','','','','','',30,100]
  ];
  sample.forEach(function(q, idx) {
    qSh.appendRow([
      setId, name, idx + 1,
      q[0], q[1], q[2],
      q[3], q[4], q[5], q[6],
      q[7], q[8],
      q[9], q[10]
    ]);
  });
  return 'เพิ่มข้อมูลตัวอย่างสำเร็จ';
}

// เพิ่ม header ชีตอัตโนมัติ (รันครั้งเดียวหลัง deploy)
function setupSheets() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  function ensureSheet(name, headers) {
    var sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    if (sh.getLastRow() === 0) sh.appendRow(headers);
    return sh;
  }

  ensureSheet(SHEET_PLAYERS, [
    'Timestamp','GameID','PlayerID','FullName','Nickname',
    'StudentID','Program','Year','Province','Expectation','TotalScore'
  ]);
  ensureSheet(SHEET_QUESTIONS, [
    'QuestionSetID','QuestionSetName','QuestionNumber','QuestionType',
    'Question','ImageURL','OptionA','OptionB','OptionC','OptionD',
    'CorrectAnswer','Explanation','TimeLimit','Points'
  ]);
  ensureSheet(SHEET_ANSWERS, [
    'Timestamp','GameID','PlayerID','Nickname','StudentID',
    'QuestionSetName','QuestionNumber','Question',
    'Answer','CorrectAnswer','IsCorrect','ResponseTime','Score'
  ]);
  ensureSheet(SHEET_SESSIONS, [
    'GameID','GamePIN','QuestionSetID','QuestionSetName',
    'StartTime','EndTime','TotalPlayers','TotalQuestions',
    'CurrentQuestionNumber','Status'
  ]);
  ensureSheet(SHEET_AQUATYPE, [
    'Timestamp','FullName','Nickname','StudentID','Program','Year','Province',
    'TopType','TopTypeName','TopTypeGroup','SecondType','SecondTypeName','SecondTypeGroup',
    'CULTScore','TECHScore','PRODScore','BIOSScore','BIZZScore','Answers'
  ]);

  return 'ตั้งค่าชีตสำเร็จ';
}

// ============================================================
// Aqua Type Test
// ============================================================
function submitAquaType(data) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sh = ss.getSheetByName(SHEET_AQUATYPE);
  if (!sh) {
    sh = ss.insertSheet(SHEET_AQUATYPE);
    sh.appendRow([
      'Timestamp','FullName','Nickname','StudentID','Program','Year','Province',
      'TopType','TopTypeName','TopTypeGroup','SecondType','SecondTypeName','SecondTypeGroup',
      'CULTScore','TECHScore','PRODScore','BIOSScore','BIZZScore','Answers'
    ]);
  }
  sh.appendRow([
    new Date().toISOString(),
    data.fullName        || '',
    data.nickname        || '',
    data.studentId       || '',
    data.program         || '',
    data.year            || '',
    data.province        || '',
    data.topType         || '',
    data.topTypeName     || '',
    data.topTypeGroup    || '',
    data.secondType      || '',
    data.secondTypeName  || '',
    data.secondTypeGroup || '',
    data.cultScore  || 0,
    data.techScore  || 0,
    data.prodScore  || 0,
    data.biosScore  || 0,
    data.bizzScore  || 0,
    data.answers    || '',
  ]);
  return { success: true };
}

function getAquaTypeStats() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sh = ss.getSheetByName(SHEET_AQUATYPE);
  var stats = { CULT:0, TECH:0, PROD:0, BIOS:0, BIZZ:0 };
  var total = 0;
  if (sh && sh.getLastRow() > 1) {
    // TopType อยู่ที่ column 8 (index 7, 0-based)
    var rows = sh.getRange(2, 1, sh.getLastRow() - 1, 8).getValues();
    rows.forEach(function(r) {
      var key = String(r[7]);
      if (stats.hasOwnProperty(key)) { stats[key]++; total++; }
    });
  }
  return { success: true, stats: stats, total: total };
}
