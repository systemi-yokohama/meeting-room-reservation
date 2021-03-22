/* global GmailApp, SpreadsheetApp */

'use strict'

const SPREADSHEET_ID = '1hxE7cBpoPz3o5o2EJtExOCUFH5SFDoOiG-vW1MrJByo'

// eslint-disable-next-line no-unused-vars
function debug (str) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID)
  const s = ss.getSheetByName('デバッグログ')
  s.appendRow([new Date().toLocaleString(), str])
}

function getMailAddresses () {
  const s = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('送信アドレス') // シート取得
  return s.getDataRange().getValues() // 項目・連絡先が入力されているセルの取得
}

// eslint-disable-next-line no-unused-vars
function submitForm (e) {
  let mailAddress = ''
  let content = ''
  let mailSubject = ''

  const mailAddresses = getMailAddresses()

  // フォーム入力内容の取得
  const itemResponses = e.response.getItemResponses()
  for (const itemResponse of itemResponses) {
    const title = itemResponse.getItem().getTitle()
    const answer = itemResponse.getResponse()
    content = `${content}[${title}]\n${answer}\n\n`
    // 送信先メールアドレスを解決する
    if (title === '相談内容を選択して下さい') {
      for (const mailAddressPair of mailAddresses) {
        if (answer === mailAddressPair[0]) {
          mailAddress = mailAddressPair[1]
          mailSubject = `【${answer}】に関する相談`
          break
        }
      }
    }
  }

  GmailApp.sendEmail(mailAddress, mailSubject, content)
}
