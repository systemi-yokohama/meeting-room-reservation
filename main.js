///use strict

const SPREADSHEET_ID = '1IUovXfXQZINOKe-oWs2J7Ipt0V7XcCle7StL5Zslr4o'

const debug = (str) => {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID)
  const s = ss.getSheetByName("デバッグログ")
  s.appendRow([new Date().toLocaleString(), str])
}

const postToSlack = (id, name) => {

  //スプレッドシート読み込み(月初日から月末日の予定を取得)
  const s = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(name + "予定出力")

  let cal = CalendarApp.getCalendarById(id) //カレンダーID取得
  let nowDate = new Date()
  let firstDate = new Date(nowDate.getFullYear(), nowDate.getMonth(), 1) //月初日を取得
  let endDate = new Date(nowDate.getFullYear(), nowDate.getMonth() + 6, 0) //月末日を取得
  Logger.log(firstDate)
  Logger.log(endDate)
  let events = cal.getEvents(firstDate, endDate)　//6か月のイベントを取得
  Logger.log(`events:${JSON.stringify(events)}`)

  const currentEventObject = events.reduce((acc, cur) => {
    acc[`${cur.getId()}-${cur.getStartTime().toISOString()}`] = {
      id: cur.getId(),
      title: cur.getTitle(),
      startTime: cur.getStartTime(),
      endTime: cur.getEndTime(),
      creators: cur.getCreators()
    }
    return acc
  }, {})
  Logger.log(`currentEventObject:${JSON.stringify(currentEventObject)}`)


  //スプレッドシートに入っている値を配列として全て取得
  const calendarEvents = s.getDataRange().getValues()

  const previousEventObject = calendarEvents.reduce((acc, cur) => {
    if (cur[0] === '') {
      return acc
    }
    const event = JSON.parse(cur)
    event.startTime = new Date(event.startTime)
    event.endTime = new Date(event.endTime)
    acc[`${event.id}-${event.startTime.toISOString()}`] = event
    return acc
  }, {})
  Logger.log(`previousEventObject:${JSON.stringify(previousEventObject)}`)

  const addedEvents = [] // 増えたイベントを保持する配列
  const changedEvents = [] // 変更されたイベントを保持する配列
  const removedEvents = [] // 削除されたイベントを保持する配列

  Object.keys(currentEventObject).forEach(key => {
    const currentEvent = currentEventObject[key]
    const previousEvent = previousEventObject[key]
    if (previousEvent) {
      for (const key of Object.keys(currentEvent)) {
        if (currentEvent[key] !== previousEvent[key]) {
          if (typeof currentEvent[key].getTime === 'function') {
            if (currentEvent[key].getTime() !== previousEvent[key].getTime()) {
              Logger.log(`key=${key}, ${currentEvent[key]}, ${previousEvent[key]}`)
              changedEvents.push(currentEvent)
              break
            }
          } else if (Array.isArray(currentEvent[key])) {
            if (JSON.stringify(currentEvent[key]) !== JSON.stringify(previousEvent[key])) {
              Logger.log(`array key=${key}, ${currentEvent[key]}, ${previousEvent[key]}`)
              changedEvents.push(currentEvent)
              break
            }
          } else {
            Logger.log(`key=${key}, '${currentEvent[key]}', '${previousEvent[key]}'`)
            changedEvents.push(currentEvent)
            break
          }
        }
      }
    } else {
      addedEvents.push(currentEvent)
    }
  })

  Object.keys(previousEventObject).forEach(key => {

    const currentEvent = currentEventObject[key]
    if (!currentEvent) {
      const previousEvent = previousEventObject[key]
      removedEvents.push(previousEvent)
    }
  })

  Logger.log(`addedEvents:${JSON.stringify(addedEvents)}`)
  Logger.log(`changedEvents:${JSON.stringify(changedEvents)}`)
  Logger.log(`removedEvents:${JSON.stringify(removedEvents)}`)

  // //日付表示の変換
  const _MMdd = (nowDate) => Utilities.formatDate(nowDate, 'JST', 'yyyy-MM-dd')
  const _HHmm = (nowDate) => Utilities.formatDate(nowDate, 'JST', 'HH:mm')

  //SlackのwebhookURLを指定
  let url = "https://hooks.slack.com/services/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" //江種
  // let url = "https://hooks.slack.com/services/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" //テスト

  let text = ''


  // まず、削除されたイベントを見る
  for (let i = 0; i < removedEvents.length; i++) {
    text += `【${removedEvents[i].title}】が削除されました\n`
  }

  const hasCreatorNameInTitle = title => /.+[（(][^（()]+[）)]$/.test(title)

  // 次に、追加されたイベントを見る
  if (addedEvents.length !== 0) {
    text += text.length === 0 ? '' : '\n'
    text += ':calendar:追加された予定\n'
    for (let i = 0; i < addedEvents.length; i++) {
      // const day = addedEvents[i].startTime
      const startTime = addedEvents[i].startTime
      const endTime = addedEvents[i].endTime
      const title = addedEvents[i].title
      const creator = hasCreatorNameInTitle(title) ? '' : `（${addedEvents[i].creators}）`
      text += ` ${_MMdd(startTime)}-${_HHmm(endTime)} ${title}${creator}\n`
    }
  }

  // 最後に、変更されたイベントを見る
  if (changedEvents.length !== 0) {
    text += text.length === 0 ? '' : '\n'
    text += ':calendar:変更された予定\n'
    for (let i = 0; i < changedEvents.length; i++) {
      // const day = 
      const startTime = changedEvents[i].startTime
      const endTime = changedEvents[i].endTime
      const title = changedEvents[i].title
      const creator = hasCreatorNameInTitle(title) ? '' : `（${changedEvents[i].creators}）`
      text += ` ${_MMdd(startTime)}-${_HHmm(endTime)} ${title}${creator}\n`
    }

  }
  let roomName = name + "予約\n"
  let data = { "username": "Googlecalendar-Bot", "text": roomName + text, "icon_emoji": ":spiral_calendar_pad: " }
  let payload = JSON.stringify(data)
  let options = {
    "method": "POST",
    "contentType": "application/json",
    "payload": payload
  }
  UrlFetchApp.fetch(url, options)

  //処理終了後、スプレッドシートをクリアし、最新の予定を記録
  s.clear()
  s.getRange(1, 1, events.length).setValues(Object.keys(currentEventObject).map(key => [JSON.stringify(currentEventObject[key])]))
}

const getMeetingRoomName = id => {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName("Calender_ID")
  const values = ss.getDataRange().getValues()
  for (let i = 0; i < values.length; i++) {
    if (values[i][0] === id) {
      return values[i][1]
    }
  }
  return null
}

const onCalendarEventUpdated = e => {
  const id = e.calendarId
  debug(id)
  const name = getMeetingRoomName(id)
  debug(name)

  postToSlack(id, name)
}
