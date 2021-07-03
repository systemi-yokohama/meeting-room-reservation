/* global Calendar, CalendarApp, LockService, Logger, PropertiesService, SpreadsheetApp, UrlFetchApp, Utilities */

'use strict'

const SPREADSHEET_ID = '1RkFxDI5wWxlZTxC8bLkR6DRHXYEvdk5jIwE64rGzYXE'

const postToSlack = (id, name) => {
  // スプレッドシート読み込み(月初日から月末日の予定を取得)
  const s = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(`${name}予定出力`)

  const cal = CalendarApp.getCalendarById(id) // カレンダーID取得
  const nowDate = new Date()
  const firstDate = new Date(nowDate.getFullYear(), nowDate.getMonth(), 1) // 月初日を取得
  const endDate = new Date(nowDate.getFullYear(), nowDate.getMonth() + 6, 0) // 月末日を取得
  Logger.log(firstDate)
  Logger.log(endDate)
  const events = cal.getEvents(firstDate, endDate) // 6か月のイベントを取得
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

  // スプレッドシートに入っている値を配列として全て取得
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

  // SlackのwebhookURLを指定
  const url = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Slack Incoming Webhook').getRange('A1').getValue()
  let text = ''

  // まず、削除されたイベントを見る
  for (let i = 0; i < removedEvents.length; i++) {
    text += `${_MMdd(removedEvents[i].startTime)} の【${removedEvents[i].title}】が削除されました。\n`
  }

  const hasCreatorNameInTitle = title => /.+[（(][^（()]+[）)]$/.test(title)

  // 次に、追加されたイベントを見る
  if (addedEvents.length !== 0) {
    text += text.length === 0 ? '' : '\n'
    text += '・追加された予定\n'
    for (let i = 0; i < addedEvents.length; i++) {
      // const day = addedEvents[i].startTime
      const startTime = addedEvents[i].startTime
      const endTime = addedEvents[i].endTime
      const title = addedEvents[i].title
      const creator = hasCreatorNameInTitle(title) ? '' : `（${addedEvents[i].creators}）`
      text += ` ${_MMdd(startTime)} ${_HHmm(startTime)}-${_HHmm(endTime)} ${title}${creator}\n`
    }
  }

  // 最後に、変更されたイベントを見る
  if (changedEvents.length !== 0) {
    text += text.length === 0 ? '' : '\n'
    text += '・変更された予定\n'
    for (let i = 0; i < changedEvents.length; i++) {
      // const day =
      const startTime = changedEvents[i].startTime
      const endTime = changedEvents[i].endTime
      const title = changedEvents[i].title
      const creator = hasCreatorNameInTitle(title) ? '' : `（${changedEvents[i].creators}）`
      text += ` ${_MMdd(startTime)} ${_HHmm(startTime)}-${_HHmm(endTime)} ${title}${creator}\n`
    }
  }

  // 予約枠を使用した際にどの配列にも入らないが、イベントが更新された判定になるため判定を追加(根本的な解決にはならない)
  if (addedEvents.length === 0 && changedEvents.length === 0 && removedEvents.length === 0) {
    return null
  }

  const roomName = `:calendar:${name}予約\n`
  const data = { username: 'Googlecalendar-Bot', text: roomName + text, icon_emoji: ':spiral_calendar_pad: ' }
  const payload = JSON.stringify(data)
  const options = {
    method: 'POST',
    contentType: 'application/json',
    payload: payload
  }
  UrlFetchApp.fetch(url, options)

  // 処理終了後、スプレッドシートをクリアし、最新の予定を記録
  s.clear()
  if (events.length !== 0) {
    s.getRange(1, 1, events.length).setValues(Object.keys(currentEventObject).map(key => [JSON.stringify(currentEventObject[key])]))
  }
}

const getMeetingRoomName = id => {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('カレンダーID')
  const values = ss.getDataRange().getValues()
  for (let i = 0; i < values.length; i++) {
    if (values[i][1] === id) {
      return values[i][0]
    }
  }
  return null
}

// eslint-disable-next-line no-unused-vars
const onCalendarEventUpdated = e => {
  const id = e.calendarId
  const name = getMeetingRoomName(id)

  const properties = PropertiesService.getScriptProperties()
  const nextSyncToken = properties.getProperty('syncToken')
  const optionalArgs = {
    syncToken: nextSyncToken
  }
  const events = Calendar.Events.list(id, optionalArgs)
  const ssETag = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('etag保存用')

  /**
   * 最新の予定のetagをスプレッドシートに保存(10個まで)
   *
   * @param {string} etag このイベントの ETag
   * @returns {boolean} 過去 10 回のイベントに同一の ETag が含まれていたか否か
   */
  const updateETag = (etag) => {
    const lock = LockService.getScriptLock()
    while (!lock.tryLock(10 * 1000)) {
      // Nothing to do
    }
    const values = ssETag.getDataRange().getValues()
    // 1列目目の1行目から9行目を、1列目の2行目へ移動させる
    ssETag.getRange(1, 1, 9, 1).moveTo(ssETag.getRange(2, 1))
    // 1列目の1行目に最新イベント固有のetagをセットする
    ssETag.getRange(1, 1).setValue([etag])
    lock.releaseLock()
    return values.flatMap(v => v).includes(etag)
  }

  // 予定イベントのetagが前回と違う場合のみslack通知を実行
  // Google カレンダーから会議室を追加して予定を作成すると同一イベントが複数回通知されることがあるが ETag が同一であるため 2 回目以降を処理しないようにする
  updateETag(events.etag) || postToSlack(id, name)
}

// デバッグ
// eslint-disable-next-line no-unused-vars
const debug = (events) => {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('デバッグ')
  ss.getRange('A1').setValue(events)
}
