function doPost (e){
  const str = { text:e.parameter.text }.text

  if (!str.indexOf('list')){
    b = /(?<command>[^ ]+)\s+(?<roomName>[^ ]+)/
    c = /(?<command>[^ ]+)\s+(?<roomName>[^ ]+)\s+(?<day>(\d{4}-\d{2}))/
  
    if (str.match(c)){
      const strGroup = str.match(c).groups

      if (isDate(strGroup.day) == 0) return ContentService.createTextOutput(str + '\n入力した日付を確認して下さい。')
      
      return ContentService.createTextOutput(getCalendarEvents(strGroup,1))
    }else if (str.match(b)){
      const strGroup = str.match(b).groups
      return ContentService.createTextOutput(getCalendarEvents(strGroup,0))
    }else{
      return ContentService.createTextOutput(roomList())
    }
  
  }else if (!str.indexOf('add')){
    d = /(?<command>[^ ]+)($|\s+)((?<roomName>[^ ]+))($|\s+(?<day>(\d{4}-\d{2}-\d{2})))($|\s+(?<startTime>\d{1,2}:\d{2})-(?<endTime>\d{1,2}:\d{2}))\s+(("(?<title1>[^"]+)")|(?<title2>[^ ]+))\s+(((?<name1>\([^"]+\)))|(?<name2>[^ ]+))/
    f = /(?<command>[^ ]+)($|\s+)((?<roomName>[^ ]+))($|\s+(?<day>(\d{4}-\d{2}-\d{2})))($|\s+(?<startTime>\d{1,2}:\d{2})-(?<endTime>\d{1,2}:\d{2}))\s+(("(?<title1>[^"]+)")|(?<title2>[^ ]+))/

    if (str.match(d)){
      const strGroup = str.match(d).groups
      const title = strGroup.title1 || strGroup.title2
      const name = strGroup.name1 || strGroup.name2
      strGroup.title = title
      strGroup.name = name
      if (isDate(strGroup.day) == 0) return ContentService.createTextOutput(str + '\n入力した日付を確認して下さい。')
      if (isTime(strGroup.startTime,strGroup.endTime) == 0) return ContentService.createTextOutput(str + '\n入力した時刻を確認して下さい。')
           
      return addRoom(strGroup)

    }else if (str.match(f)){
      const strGroup = str.match(f).groups
      const title = strGroup.title1 || strGroup.title2
      strGroup.title = title
      if (isDate(strGroup.day) == 0) return ContentService.createTextOutput(str + '\n入力した日付を確認して下さい。')
      if (isTime(strGroup.startTime,strGroup.endTime) == 0) return ContentService.createTextOutput(str + '\n入力した時刻を確認して下さい。')
           
      return addRoom(strGroup)
    }

  }else if (!str.indexOf('help')){
    return ContentService.createTextOutput('<https://systemi.backlog.com/wiki/GENINFO/%E4%BC%9A%E8%AD%B0%E5%AE%A4%E3%81%AE%E4%BA%88%E7%B4%84%E6%96%B9%E6%B3%95|会議室予約方法>')
  
  }else if (str == ''){
    return ContentService.createTextOutput('コマンドが入力されていません。')
  }else{
    return ContentService.createTextOutput(str + '\nコマンドを正しく入力してください。')
  }
}


//スプレッドシートからカレンダーIDを取得する
function getRoomId (roomName) {
  ss = SpreadsheetApp.getActiveSpreadsheet()
  s = ss.getSheetByName('カレンダーID')

  datas = s.getDataRange().getValues()
  
  for (const data of datas) {
    if (data.indexOf(roomName) >= 0) {
      return data
    }
  }
  return 0
 
}

//イベントを削除する処理
//「予定削除」のリンクを押したらここで処理される
function doGet (e) {

  const calendar = CalendarApp.getCalendarById(e.parameter.calendarId) //パラメーターのcalenderIdからカレンダーを指定する
  const event= calendar.getEventById(e.parameter.eventId) //パラメーターのeventIdから予定を指定する
  
  event.deleteEvent()

  return ContentService.createTextOutput('予定を削除しました。')
  
}


//カレンダーに予定を追加する
function addRoom (strGroup) {

  //スプレッドシートからカレンダーIDを取得する
  const roomdata = getRoomId(strGroup.roomName)
  if (roomdata == 0) return ContentService.createTextOutput('入力した部屋名を確認して下さい。')
  const calendarId = roomdata[1]
  strGroup.roomName = roomdata[2]

  const calendar = CalendarApp.getCalendarById(calendarId)
  let title
  if (strGroup.name == undefined){
    title = strGroup.title
  }else{
    title = strGroup.title + ' ' +strGroup.name
  }
  const startTime = new Date(strGroup.day + ' ' + strGroup.startTime)
  const endTime = new Date(strGroup.day + ' ' + strGroup.endTime)
  calendar.createEvent(title, startTime, endTime)

  return ContentService.createTextOutput('予約内容　'+strGroup.roomName+' '+strGroup.day+' '+strGroup.startTime+'-'+strGroup.endTime)
}


//カレンダーから予定を取得する
function getCalendarEvents (strGroup,setday=0) {

  //スプレッドシートからカレンダーIDを取得する
  const roomdata = getRoomId(strGroup.roomName)
  if (roomdata == 0) return ('入力した部屋名を確認して下さい。')
  const calendarId = roomdata[1]
  strGroup.roomName = roomdata[2]

  const calendar = CalendarApp.getCalendarById(calendarId)
  if (setday == 0){
    const today = new Date() //取得された日にち
    const startTime = new Date(today.getFullYear(), today.getMonth(), 1,00,00,00) //取得された月の初めの時間
    const endTime = new Date(today.getFullYear(), today.getMonth(), 31,24,00,00) //取得された月の終わりの時間
    const events = calendar.getEvents(startTime, endTime)

    return getList(events,calendarId,strGroup.roomName)

  }else{
    const today = new Date(strGroup.day + '-01 00:00:00')
    const startTime = new Date(today.getFullYear(), today.getMonth(), 1,00,00,00) //指定された月の初めの時間
    const endTime = new Date(today.getFullYear(), today.getMonth(), 31,24,00,00) //指定された月の終わりの時間
    const events = calendar.getEvents(startTime, endTime)
    
    return getList(events,calendarId,strGroup.roomName)
  }
}

//listをメッセージにして返す
function getList (events,calenderId,roomName) {
  let reserveList = roomName +'の予約状況\n'
    for(const event of events){
      const title = event.getTitle()
      const event_ID = event.getId()
      const start =　Utilities.formatDate( event.getStartTime(), 'JST', 'yyyy-MM-dd HH:mm')
      const end = Utilities.formatDate(event.getEndTime(), 'JST', 'HH:mm')
      const b = ( start +'-'+ end +' '+ title + '　<https://script.google.com/a/systemi.co.jp/macros/s/AKfycby7qfJI5wXCXLV3QQPMv85C-ddctFnPnS81o3vQ/exec?roomId=' +calenderId + '&eventId='+event_ID+'|予定削除>\n')
      reserveList += b
      
    }
    return reserveList
}

function roomList () {
  ss = SpreadsheetApp.getActiveSpreadsheet()
  s = ss.getSheetByName('カレンダーID')

  datas = s.getDataRange().getValues()

  let roomlist = '予約可能な会議室一覧\n'

  for (let data of datas){
    roomlist += '・' + data[0] + '\n'
  }

  return roomlist
}

function reservation (reserve) {
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  const s = ss.getSheetByName("予約管理")
  s.appendRow(reserve)
}

function isDate (date) {
  //YYYY-mm の判定　この場合mmの範囲のみの判定
  if (date.match(/\d{4}-\d{2}-\d{2}/)){
    let mm = date.match(/(?<year>\d{4})-(?<month>\d{2})-(?<day>\d{2})/).groups.month
    let dd = date.match(/(?<year>\d{4})-(?<month>\d{2})-(?<day>\d{2})/).groups.day

    mm = mm - 0
    dd = dd - 0

    if (mm < 1 || 12 < mm || dd < 1 || 31 < dd){
      return 0
    }else{
      console.log(date)
    }
  // mm の判定
  }else if (date.match(/\d{4}-\d{2}/)){
    let mm = date.match(/(?<year>\d{4})-(?<month>\d{2})/).groups.month
    
    mm = mm - 0

    if (mm < 1 || 12 < mm){
      return 0
    }else{
      return 1
    }
  
  }
}

function isTime (starttime,endtime) {
  
  let start = starttime.match(/(?<hour>\d{1,2}):(?<minute>\d{2})/).groups
  let end = endtime.match(/(?<hour>\d{1,2}):(?<minute>\d{2})/).groups

  s_hh = start.hour - 0
  s_mm = start.minute - 0
  e_hh = end.hour - 0
  e_mm = end.minute - 0

  if (s_hh < 0 || 24 < s_hh || e_hh < 0 || 24 < e_hh || s_hh > e_hh){
    return 0

  }else if (s_mm < 0 || 59 < s_mm || e_mm < 0 || 59 < e_mm){
    return 0

  }else if (s_hh == e_hh && s_mm >= e_mm){
    return 0
  }else{
    return 1
  }
}

