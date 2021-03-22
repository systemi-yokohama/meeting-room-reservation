//use strict
const debug = (str) => {
  const ss = SpreadsheetApp.openById("1IUovXfXQZINOKe-oWs2J7Ipt0V7XcCle7StL5Zslr4o")
  const s = ss.getSheetByName("デバッグログ")
  s.appendRow([new Date().toLocaleString(), str])
}

let nowDate = new Date()
let firstDate = new Date(nowDate.getFullYear(), nowDate.getMonth(), 1) //月初日を取得
let endDate = new Date(nowDate.getFullYear(), nowDate.getMonth() + 6, 0) //月末日を取得
Logger.log(firstDate)
Logger.log(endDate)

const hoge = (e) => debug(JSON.stringify(e))

//日付表示の変換
const _MMdd = (date) => Utilities.formatDate(date, 'JST', 'M/d (E)')
const _HHmm = (hours) => Utilities.formatDate(hours, 'JST', 'HH:mm')

Calendar = (id, name) => {

  //スプレッドシート読み込み(月初日から月末日の予定を取得)
  const s = SpreadsheetApp.openById('1IUovXfXQZINOKe-oWs2J7Ipt0V7XcCle7StL5Zslr4o').getSheetByName(name + "予定出力")

  let cal = CalendarApp.getCalendarById(id) //カレンダーID取得
  let monthLater = new Date()
  monthLater.setMonth(firstDate.getMonth() + 1, 0)　//１か月の日付を取得
  let events = cal.getEvents(firstDate, monthLater)　//１か月のイベントを取得
  let arrDate = []//日付用の配列
  let arrTitle = []//イベント名の配列
  let arrUpDate = []//更新日時の配列
  let arrStart_time = [] //イベントの開始時刻の配列
  let arrEnd_time = [] //イベント終了時刻の配列
  let arrCreators = [] //予定作成者の配列

  //前回記録した予定の最終行を取得
  const columns1 = s.getRange(1, 1).getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow();
  Logger.log(columns1)

  //イベントの日付、タイトル、更新日時、開始時刻、終了時刻、予定作成者を取得
  for (let i = 0; i < events.length; i++) {
    arrDate.push(events[i].getStartTime()),
      arrTitle.push(events[i].getTitle()),
      arrUpDate.push(events[i].getLastUpdated().getTime()),
      arrStart_time.push(events[i].getStartTime()),
      arrEnd_time.push(events[i].getEndTime()),
      arrCreators.push(events[i].getCreators())
  }

  //最新のイベントを取得
  let maxDay = Math.max.apply(null, arrUpDate)
  let index = arrUpDate.indexOf(maxDay)

  //前回記録した予定の数と今回取得したイベント数を比較
  if (events.length < columns1) {
    events[index].deleteEvent()
  } else {
    //SlackのwebhookURLを指定
    let postMsg = name + "予約\n" + _MMdd(arrDate[index]) + " " + _HHmm(arrStart_time[index]) + "-" + _HHmm(arrEnd_time[index]) + " " + arrTitle[index] + " " + arrCreators[index]

    let url = "https://hooks.slack.com/services/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"
    //渡すデータを指定する
    let data = { /*"channel" : ch,*/ "username": "Googlecalendar-Bot", "text": postMsg, "icon_emoji": ":spiral_calendar_pad: " }
    let payload = JSON.stringify(data)
    let options = {
      "method": "POST",
      "contentType": "application/json",
      "payload": payload
    }
    UrlFetchApp.fetch(url, options)
  }

  //処理終了後、スプレッドシートをクリアし、最新の予定を記録
  s.clear()
  for (let i = 0; i < events.length; i++) {
    s.appendRow(
      [
        events[i].getStartTime(),
        events[i].getTitle(),
        events[i].getLastUpdated().getTime(),
        events[i].getStartTime(),
        events[i].getEndTime(),
      ]
    )
  }
  const columns2 = s.getRange(1, 1).getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow();
  Logger.log(columns2)
}


const ss = SpreadsheetApp.openById('1IUovXfXQZINOKe-oWs2J7Ipt0V7XcCle7StL5Zslr4o').getSheetByName("Calender_ID")
const values = ss.getDataRange().getValues()

Logger.log(values)
// for (value of values){
//   if(events === value){
//     Calendar([value][value])
//   }
// }
Calendar('y.egusa@systemi.co.jp', "応接室")



// let maxDay = Math.max.apply(null, arrUpDate)
// let index = arrUpDate.indexOf(maxDay)

// Logger.log(s1.getMaxColumns())

// //SlackのwebhookURLを指定
// postMsg = "会議室予約\n" + _MMdd(arrDate[index]) + " " + _HHmm(arrStart_time[index]) + "-" + _HHmm(arrEnd_time[index]) + " " + arrTitle[index] + " " + arrCreators[index]

// let url = "https://hooks.slack.com/services/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"
// //渡すデータを指定する
// let data = { /*"channel" : ch,*/ "username": "Googlecalendar-Bot", "text": postMsg, "icon_emoji": ":spiral_calendar_pad: " }
// let payload = JSON.stringify(data)
// let options = {
//   "method": "POST",
//   "contentType": "application/json",
//   "payload": payload
// }
// UrlFetchApp.fetch(url, options)




// for (let i = 0; i < ids.length; i++) {
//   if (idsevents === 'y.egusa@systemi.co.jp') {
//     Calendar(idsevents)
//   }
//   else {
//     Calendar(idsevents)
//   }
// }


// Calendar("y.egusa@systmei.co.jp")

// function day() {

//   // 曜日の配列
//   const week_list = ['日', '月', '火', '水', '木', '金', '土']

//   // 曜日を表す数値
//   const weekNum = nowDate.getDay()

//   // 曜日取得
//   const week = week_list[weekNum]
//   return week
// }

// function sampleCreators(calender_id) {
//   let calendarId = CalendarApp.getCalendarById(calender_id)
//   let date = new Date()
//   let monthLater = new Date()
//   let events =CalendarApp.getCalendarById(calendarId).getEvents(date, monthLater)
//   monthLater.setMonth(nowDate.getMonth() + 1000)
//   for (let i in events) {
//     let event = events
//     let creators = event.getCreators()
//     Logger.log(creators)
//   }
//   return creators
// }


// const ids = getIds()
// const ids = [
//   ['江種','y.egusa@systemi.co.jp'],
//   ['会議室','c_59fojvt0b8c2tq4pcm1i917di8@group.calendar.google.com'],
//   ['応接室','c_5i9japei0bkg8vf589t0hjjfd4@group.calendar.google.com'],
//   ['桜木町シティビル会議スペース','c_56e76201ttkorffuhel1ml4prs@group.calendar.google.com'],
//   ['テスト用','c_67e5qugtvehjagh1gjpg29gjbs@group.calendar.google.com']
// ]

