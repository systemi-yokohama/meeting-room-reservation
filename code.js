/* global CalendarApp, ContentService, SpreadsheetApp, Utilities */

'use strict'

// eslint-disable-next-line no-unused-vars
function doPost (e) {
  const str = { text: e.parameter.text }.text
  const hyfun = '[ー‐−―－-]'

  if (!str.indexOf('list')) {
    const b = /(?<command>[^\s]+)\s+(?<roomName>[^\s]+)/
    const c = new RegExp(`(?<command>[^\\s]+)\\s+(?<roomName>[^\\s]+)\\s+(?<day>[0-9０-９]{4}${hyfun}[0-9０-９]{2})`)

    if (str.match(c)) {
      const strGroup = str.match(c).groups
      strGroup.day = regexp(strGroup.day)

      try {
        isDate(strGroup.day)
      } catch (error) {
        return ContentService.createTextOutput(error.message)
      }

      return ContentService.createTextOutput(getCalendarEvents(strGroup, 1))
    } else if (str.match(b)) {
      const strGroup = str.match(b).groups
      return ContentService.createTextOutput(getCalendarEvents(strGroup, 0))
    } else {
      return ContentService.createTextOutput(roomList())
    }
  } else if (!str.indexOf('add')) {
    const time = '[0-9０-９]{1,2}[:：][0-9０-９]{2}'

    const d = new RegExp(`(?<command>[^\\s]+)($|\\s+)((?<roomName>[^\\s]+))($|\\s+(?<day>([0-9０-９]{4}${hyfun}[0-9０-９]{2}${hyfun}[0-9０-９]{2})))($|\\s+(?<startTime>${time})${hyfun}(?<endTime>${time}))\\s+((["“”](?<title1>[^"“”]+)["“”])|(?<title2>[^\\s]+))\\s+(((?<name1>[(（][^["“”]]+[)）]))|(?<name2>[^\\s]+))`)

    const f = new RegExp(`(?<command>[^\\s]+)($|\\s+)((?<roomName>[^\\s]+))($|\\s+(?<day>([0-9０-９]{4}${hyfun}[0-9０-９]{2}${hyfun}[0-9０-９]{2})))($|\\s+(?<startTime>${time})${hyfun}(?<endTime>${time}))\\s+((["“”](?<title1>[^"“”]+)["“”])|(?<title2>[^\\s]+))`)

    const h = (`(?<command>[^\\s]+)($|\\s+)((?<roomName>[^\\s]+))($|\\s+(?<day>([0-9０-９]{4}${hyfun}[0-9０-９]{2}${hyfun}[0-9０-９]{2})))`)

    if (str.match(d)) {
      const strGroup = str.match(d).groups
      const title = strGroup.title1 || strGroup.title2
      const name = strGroup.name1 || strGroup.name2
      strGroup.title = title
      strGroup.name = name

      strGroup.startTime = regexp(strGroup.startTime)
      strGroup.endTime = regexp(strGroup.endTime)
      strGroup.day = regexp(strGroup.day)

      try {
        isDate(strGroup.day)
      } catch (error) {
        return ContentService.createTextOutput(error.message)
      }
      try {
        isTime(strGroup.startTime, strGroup.endTime)
      } catch (error) {
        return ContentService.createTextOutput(error.message)
      }

      return addRoom(strGroup)
    } else if (str.match(f)) {
      const strGroup = str.match(f).groups
      const title = strGroup.title1 || strGroup.title2
      strGroup.title = title

      strGroup.startTime = regexp(strGroup.startTime)
      strGroup.endTime = regexp(strGroup.endTime)
      strGroup.day = regexp(strGroup.day)

      try {
        isDate(strGroup.day)
      } catch (error) {
        return ContentService.createTextOutput(error.message)
      }
      try {
        isTime(strGroup.startTime, strGroup.endTime)
      } catch (error) {
        return ContentService.createTextOutput(error.message)
      }

      return addRoom(strGroup)
    } else if (str.match(h)) {
      return ContentService.createTextOutput(str + '\n予定を追加する場合は日時、タイトルを入力してください。')
    }
  } else if (!str.indexOf('help')) {
    return ContentService.createTextOutput('<https://systemi.backlog.com/wiki/GENINFO/%E4%BC%9A%E8%AD%B0%E5%AE%A4%E3%81%AE%E4%BA%88%E7%B4%84%E6%96%B9%E6%B3%95|会議室予約方法>')
  } else if (str === '') {
    return ContentService.createTextOutput('コマンドが入力されていません。')
  } else {
    return ContentService.createTextOutput(`\`${str}\`\nコマンドを正しく入力してください。`)
  }
}

function regexp (str) {
  return str.replace(/[０-９：―ー－]/g, function (s) {
    return String.fromCharCode(s.charCodeAt(0) - 0xFEE0)
  })
}

// スプレッドシートからカレンダーIDを取得する
function getRoomId (roomName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  const s = ss.getSheetByName('カレンダーID')

  const datas = s.getDataRange().getValues()

  for (const data of datas) {
    if (data.indexOf(roomName) >= 0) {
      return data
    }
  }
  return 0
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
function addRoom (strGroup) {
  // スプレッドシートからカレンダーIDを取得する
  const roomdata = getRoomId(strGroup.roomName)
  if (roomdata === 0) return ContentService.createTextOutput('入力した部屋名を確認して下さい。')
  const calendarId = roomdata[1]
  strGroup.roomName = roomdata[2]

  const calendar = CalendarApp.getCalendarById(calendarId)
  let title
  if (strGroup.name === undefined) {
    title = strGroup.title
  } else {
    title = strGroup.title + ' ' + strGroup.name
  }
  const startTime = new Date(strGroup.day + ' ' + strGroup.startTime)
  const endTime = new Date(strGroup.day + ' ' + strGroup.endTime)
  calendar.createEvent(title, startTime, endTime)

  return ContentService.createTextOutput('予約内容　' + strGroup.roomName + ' ' + strGroup.day + ' ' + strGroup.startTime + '-' + strGroup.endTime)
}

// カレンダーから予定を取得する
function getCalendarEvents (strGroup, setday = 0) {
  // スプレッドシートからカレンダーIDを取得する
  const roomdata = getRoomId(strGroup.roomName)
  if (roomdata === 0) return ('入力した部屋名を確認して下さい。')
  const calendarId = roomdata[1]
  strGroup.roomName = roomdata[2]

  const calendar = CalendarApp.getCalendarById(calendarId)
  if (setday === 0) {
    const today = new Date() // 取得された日にち
    const startTime = new Date(today.getFullYear(), today.getMonth(), 1, 0, 0, 0) // 取得された月の初めの時間
    const endTime = new Date(today.getFullYear(), today.getMonth(), 31, 24, 0, 0) // 取得された月の終わりの時間
    const events = calendar.getEvents(startTime, endTime)

    return getList(events, calendarId, strGroup.roomName)
  } else {
    const today = new Date(strGroup.day + '-01 00:00:00')
    const startTime = new Date(today.getFullYear(), today.getMonth(), 1, 0, 0, 0) // 指定された月の初めの時間
    const endTime = new Date(today.getFullYear(), today.getMonth(), 31, 24, 0, 0) // 指定された月の終わりの時間
    const events = calendar.getEvents(startTime, endTime)

    return getList(events, calendarId, strGroup.roomName)
  }
}

// listをメッセージにして返す
function getList (events, calenderId, roomName) {
  let reserveList = roomName + 'の予約状況\n'
  for (const event of events) {
    const title = event.getTitle()
    const eventID = event.getId()
    const start = Utilities.formatDate(event.getStartTime(), 'JST', 'yyyy-MM-dd HH:mm')
    const end = Utilities.formatDate(event.getEndTime(), 'JST', 'HH:mm')
    const b = (start + '-' + end + ' ' + title + '　<https://script.google.com/a/systemi.co.jp/macros/s/AKfycby7qfJI5wXCXLV3QQPMv85C-ddctFnPnS81o3vQ/exec?roomId=' + calenderId + '&eventId=' + eventID + '|予定削除>\n')
    reserveList += b
  }
  return reserveList
}

function roomList () {
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  const s = ss.getSheetByName('カレンダーID')

  const datas = s.getDataRange().getValues()

  let roomlist = '予約可能な会議室一覧\n'

  for (const data of datas) {
    roomlist += '・' + data[0] + '\n'
  }

  return roomlist
}

function isDate (date) {
  const hyfun = '[ー‐−―－-]'
  // YYYY-mm の判定 この場合mmの範囲のみの判定
  if (date.match(`[0-9０-９]{4}${hyfun}[0-9０-９]{2}${hyfun}[0-9０-９]{2}`)) {
    let mm = date.match(`(?<year>[0-9０-９]{4})${hyfun}(?<month>[0-9０-９]{2})${hyfun}(?<day>[0-9０-９]{2})`).groups.month
    let dd = date.match(`(?<year>[0-9０-９]{4})${hyfun}(?<month>[0-9０-９]{2})${hyfun}(?<day>[0-9０-９]{2})`).groups.day

    mm = Number(mm)
    dd = Number(dd)

    if (mm < 1 || mm > 12 || dd < 1 || dd > 31) {
      throw new Error(`\`${date}\` は正しい日付ではありません。`)
    }
  // mm の判定
  } else if (date.match(`[0-9０-９]{4}${hyfun}[0-9０-９]{2}`)) {
    let mm = date.match(`(?<year>[0-9０-９]{4})${hyfun}(?<month>[0-9０-９]{2})`).groups.month

    mm = Number(mm)

    if (mm < 1 || mm > 12) {
      throw new Error(`\`${date}\` は正しい日付ではありません。`)
    }
  }
}

function isTime (starttime, endtime) {
  const start = starttime.match(/(?<hour>[0-9０-９]{1,2})[:：](?<minute>[0-9０-９]{2})/).groups
  const end = endtime.match(/(?<hour>[0-9０-９]{1,2})[:：](?<minute>[0-9０-９]{2})/).groups

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
