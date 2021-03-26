/* global CalendarApp, ContentService, SpreadsheetApp, Utilities */

'use strict'

const CALENDAR_ICON = ':calendar:'
const hyfun = '[ー‐−―－-]'

// eslint-disable-next-line no-unused-vars
function doPost (e) {
  const str = e.parameter.text

  if (str.startsWith('list')) {
    const b = /(?<command>[^\s]+)\s+(?<roomName>[^\s]+)/
    const c = new RegExp(`(?<command>[^\\s]+)\\s+(?<roomName>[^\\s]+)\\s+(?<day>[0-9０-９]{4}${hyfun}[0-9０-９]{2})`)

    let matchB = null
    const matchC = str.match(c)
    if (matchC) {
      const strGroup = matchC.groups
      strGroup.day = normalize(strGroup.day)

      try {
        tryToCheckDate(strGroup.day)
      } catch (error) {
        return ContentService.createTextOutput(error.message)
      }

      return ContentService.createTextOutput(getCalendarEvents(strGroup.day, strGroup.roomName, false))
    } else if ((matchB = str.match(b))) {
      const strGroup = matchB.groups
      return ContentService.createTextOutput(getCalendarEvents(strGroup.day, strGroup.roomName, true))
    } else {
      return ContentService.createTextOutput(roomList())
    }
  } else if (str.startsWith('add')) {
    const time = '[0-9０-９]{1,2}[:：][0-9０-９]{2}'

    const d = new RegExp(`(?<command>[^\\s]+)($|\\s+)((?<roomName>[^\\s]+))($|\\s+(?<day>([0-9０-９]{4}${hyfun}[0-9０-９]{2}${hyfun}[0-9０-９]{2})))($|\\s+(?<startTime>${time})${hyfun}(?<endTime>${time}))\\s+((["“”](?<title1>[^"“”]+)["“”])|(?<title2>[^\\s]+))\\s+(((?<name1>[(（][^["“”]]+[)）]))|(?<name2>[^\\s]+))`)

    const f = new RegExp(`(?<command>[^\\s]+)($|\\s+)((?<roomName>[^\\s]+))($|\\s+(?<day>([0-9０-９]{4}${hyfun}[0-9０-９]{2}${hyfun}[0-9０-９]{2})))($|\\s+(?<startTime>${time})${hyfun}(?<endTime>${time}))\\s+((["“”](?<title1>[^"“”]+)["“”])|(?<title2>[^\\s]+))`)

    const h = (`(?<command>[^\\s]+)($|\\s+)((?<roomName>[^\\s]+))($|\\s+(?<day>([0-9０-９]{4}${hyfun}[0-9０-９]{2}${hyfun}[0-9０-９]{2})))`)

    let matchF = null
    const matchD = str.match(d)
    if (matchD) {
      const strGroup = matchD.groups
      strGroup.title = strGroup.title1 || strGroup.title2
      strGroup.name = strGroup.name1 || strGroup.name2

      strGroup.startTime = normalize(strGroup.startTime)
      strGroup.endTime = normalize(strGroup.endTime)
      strGroup.day = normalize(strGroup.day)

      try {
        tryToCheckDate(strGroup.day)
      } catch (error) {
        return ContentService.createTextOutput(error.message)
      }
      try {
        tryToCheckTime(strGroup.startTime, strGroup.endTime)
      } catch (error) {
        return ContentService.createTextOutput(error.message)
      }

      return addEventToRoom(strGroup)
    } else if ((matchF = str.match(f))) {
      const strGroup = matchF.groups
      strGroup.title = strGroup.title1 || strGroup.title2

      strGroup.startTime = normalize(strGroup.startTime)
      strGroup.endTime = normalize(strGroup.endTime)
      strGroup.day = normalize(strGroup.day)

      try {
        tryToCheckDate(strGroup.day)
      } catch (error) {
        return ContentService.createTextOutput(error.message)
      }
      try {
        tryToCheckTime(strGroup.startTime, strGroup.endTime)
      } catch (error) {
        return ContentService.createTextOutput(error.message)
      }

      return addEventToRoom(strGroup)
    } else if (str.match(h)) {
      return ContentService.createTextOutput(str + '\n予定を追加する場合は日時、タイトルを入力してください。')
    } else {
      return ContentService.createTextOutput(`\`${str}\`\nコマンドを正しく入力してください（エラーコード: 001）。`)
    }
  } else if (str.startsWith('help')) {
    return ContentService.createTextOutput('Backlog の<https://systemi.backlog.com/wiki/GENINFO/%E4%BC%9A%E8%AD%B0%E5%AE%A4%E3%81%AE%E4%BA%88%E7%B4%84%E6%96%B9%E6%B3%95|会議室予約方法>を参照下さい（リンクをクリックするとブラウザが開きます）。')
  } else if (str === '') {
    return ContentService.createTextOutput('コマンドが入力されていません。')
  } else {
    return ContentService.createTextOutput(`\`${str}\`\nコマンドを正しく入力してください（エラーコード: 002）。`)
  }
}

function normalize (str) {
  return str.replace(/[―ー－]/g, '-').replace(/[０-９：]/g, s => String.fromCharCode(s.charCodeAt(0) - 0xFEE0))
}

// スプレッドシートからカレンダーIDを含むレコードを取得する
function getRoomIdRecord (roomName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  const s = ss.getSheetByName('カレンダーID')

  const datas = s.getDataRange().getValues()

  for (const data of datas) {
    if (data.includes(roomName)) {
      return data
    }
  }
  return null
}

// イベントを削除する処理
// 「予定削除」のリンクを押したらここで処理される
// eslint-disable-next-line no-unused-vars
function doGet (e) {
  const calendar = CalendarApp.getCalendarById(e.parameter.calendarId) // パラメーターのcalenderIdからカレンダーを指定する
  const event = calendar.getEventById(e.parameter.eventId) // パラメーターのeventIdから予定を指定する

  event.deleteEvent()

  return ContentService.createTextOutput('予定を削除しました。')
}

// カレンダーに予定を追加する
function addEventToRoom (strGroup) {
  // スプレッドシートからカレンダーIDを取得する
  const roomdata = getRoomIdRecord(strGroup.roomName)
  if (!roomdata) return ContentService.createTextOutput('入力した部屋名を確認して下さい。')
  const calendarId = roomdata[1]
  strGroup.roomName = roomdata[2]

  const calendar = CalendarApp.getCalendarById(calendarId)
  const title = strGroup.name === undefined ? strGroup.title : strGroup.title + ' ' + strGroup.name
  const startTime = new Date(strGroup.day + ' ' + strGroup.startTime)
  const endTime = new Date(strGroup.day + ' ' + strGroup.endTime)
  calendar.createEvent(title, startTime, endTime)

  return ContentService.createTextOutput(`予約内容: ${strGroup.roomName} ${strGroup.day} ${strGroup.startTime} ${strGroup.endTime}`)
}

// カレンダーから予定を取得する
function getCalendarEvents (targetMonth, roomName, isThisMonth) {
  // スプレッドシートからカレンダーIDを取得する
  const roomdata = getRoomIdRecord(roomName)
  if (roomdata === 0) return ('入力した部屋名を確認して下さい。')
  const calendarId = roomdata[1]
  roomName = roomdata[2]

  const calendar = CalendarApp.getCalendarById(calendarId)
  if (isThisMonth) {
    const today = new Date() // 取得された日にち
    const startTime = new Date(today.getFullYear(), today.getMonth()) // 取得された月の初めの時間
    const endTime = new Date(today.getFullYear(), today.getMonth() + 1) // 取得された月の終わりの時間
    const events = calendar.getEvents(startTime, endTime)

    return getList(events, calendarId, roomName)
  } else {
    const today = new Date(targetMonth + '-01 00:00:00')
    const startTime = new Date(today.getFullYear(), today.getMonth()) // 指定された月の初めの時間
    const endTime = new Date(today.getFullYear(), today.getMonth() + 1) // 指定された月の終わりの時間
    const events = calendar.getEvents(startTime, endTime)

    return getList(events, calendarId, roomName)
  }
}

// listをメッセージにして返す
function getList (events, calenderId, roomName) {
  return events.reduce((acc, cur) => {
    const title = cur.getTitle()
    const eventID = cur.getId()
    const start = Utilities.formatDate(cur.getStartTime(), 'JST', 'yyyy-MM-dd HH:mm')
    const end = Utilities.formatDate(cur.getEndTime(), 'JST', 'HH:mm')
    const b = `${start}-${end} ${title}  <https://script.google.com/a/systemi.co.jp/macros/s/AKfycby7qfJI5wXCXLV3QQPMv85C-ddctFnPnS81o3vQ/exec?roomId=${calenderId}&eventId=${eventID}|予定削除>`
    return acc + '\n' + b
  }, `${CALENDAR_ICON}${roomName}の予約状況`)
}

function roomList () {
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  const s = ss.getSheetByName('カレンダーID')
  const datas = s.getDataRange().getValues()
  return datas.reduce((acc, cur) => `${acc}\n・${cur[0]}`, `${CALENDAR_ICON}予約可能な会議室一覧`)
}

function tryToCheckDate (date) {
  // YYYY-mm の判定 この場合mmの範囲のみの判定
  let matchYearMonth = null
  const matchYearMonthDay = date.match(/(?<year>[0-9]{4})-(?<month>[0-9]{2})-(?<day>[0-9]{2})/)
  if (matchYearMonthDay) {
    let mm = matchYearMonthDay.groups.month
    let dd = matchYearMonthDay.groups.day

    mm = Number(mm)
    dd = Number(dd)

    if (mm < 1 || mm > 12 || dd < 1 || dd > 31) { // FIXME: 日付のチェックが正確ではないが良しとする
      throw new Error(`\`${date}\` は正しい日付ではありません。`)
    }
  // mm の判定
  } else if ((matchYearMonth = date.match(/(?<year>[0-9]{4})-(?<month>[0-9]{2})/))) {
    let mm = matchYearMonth.groups.month

    mm = Number(mm)

    if (mm < 1 || mm > 12) {
      throw new Error(`\`${date}\` は正しい日付ではありません。`)
    }
  }
}

function tryToCheckTime (starttime, endtime) {
  const regex = /(?<hour>[0-9]{1,2}):(?<minute>[0-9]{2})/
  const start = starttime.match(regex).groups
  const end = endtime.match(regex).groups

  const startHh = Number(start.hour)
  const startMm = Number(start.minute)
  const endHh = Number(end.hour)
  const endMm = Number(end.minute)

  if (startHh < 0 || startHh > 24 || startMm < 0 || startMm > 59) {
    throw new Error(`\`${starttime}\` は正しい時刻ではありません。`)
  } else if (endHh < 0 || endHh > 24 || endMm < 0 || endMm > 59) {
    throw new Error(`\`${endtime}\` は正しい時刻ではありません。`)
  } else if (startHh > endHh || (startHh === endHh && startMm >= endMm)) {
    throw new Error(`開始時刻が終了時刻よりも未来に指定されています: \`${starttime}-${endtime}\``)
  }
}
