/**
 * SlackのAPIトークン。
 * @type {string}
 * @const
 * @description SlackのAPIにアクセスするために必要なトークン。このトークンは秘密情報なので、外部に漏れないように注意する必要があります。
 */
const SLACK_TOKEN = PropertiesService.getScriptProperties().getProperty('BotUserOAuthToken')


/**
 * Slackの全てのチャンネルIDを取得する関数。
 * @return {Array} チャンネルIDの配列。
 */
function getAllChannelIds() {
  const url = `https://slack.com/api/conversations.list`
  const options = {
    method: "get",
    contentType: "application/x-www-form-urlencoded",
    headers: { "Authorization": `Bearer ${SLACK_TOKEN}` },
    payload: {
      "exclude_archived": true,
      "limit": 1000,
      "types": "public_channel,private_channel,mpim,im"
    }
  }

  const channelsResponse = UrlFetchApp.fetch(url, options)
  const channelIDs = JSON.parse(channelsResponse.getContentText()).channels.map(channel => channel.id)
  // console.log(channelIDs)
  return channelIDs
}


/**
 * Slackのワークスペースでの会話において、直近1ヶ月間の利用されている絵文字のランキングを集計する関数。
 * @return {string} トップ10の絵文字とその使用回数の文字列。
 */
function collectEmojiRanking() {
  const channelIds = getAllChannelIds()
  const emojiCounts = {}

  // 現在の日時から1ヶ月前のUNIXタイムスタンプを計算
  const today = new Date()
  today.setMonth(today.getMonth() - 1)
  // console.log(today)
  const oldestTimestamp = Math.floor(today.getTime() / 1000)
  // console.log(oldestTimestamp)

  channelIds.forEach(channelId => {
    // Slack APIのURLを設定
    const url = `https://slack.com/api/conversations.history`

    // APIリクエストのオプションを設定
    const options = {
      method: "get",
      contentType: "application/x-www-form-urlencoded",
      headers: { "Authorization": `Bearer ${SLACK_TOKEN}` },
      payload: {
        "channel": channelId,
        "limit": 1000,
        "oldest": oldestTimestamp  // 1ヶ月前のタイムスタンプを指定
      }
    };

    // Slack APIからメッセージ履歴を取得
    const historyResponse = UrlFetchApp.fetch(url, options)
    const messages = JSON.parse(historyResponse.getContentText()).messages || []

    messages.forEach(message => {
      // 会話のテキスト内の絵文字を集計
      const emojisInText = (message.text || "").match(/:(\w+):/g) || []
      emojisInText.forEach(emoji => {
        if (!emojiCounts[emoji]) {
          emojiCounts[emoji] = 0
        }
        emojiCounts[emoji]++
      })

      // リアクションとして追加された絵文字を集計
      if (message.reactions) {
        message.reactions.forEach(reaction => {
          const emoji = `:${reaction.name}:`
          if (!emojiCounts[emoji]) {
            emojiCounts[emoji] = 0
          }
          // リアクションの回数（count）を絵文字の使用回数に追加
          emojiCounts[emoji] += reaction.count
        })
      }
    })
  })

  // 絵文字の使用回数を降順にソートして、トップ10の絵文字を取得
  const sortedEmojiCounts = Object.entries(emojiCounts).sort((a, b) => b[1] - a[1]).slice(0, 10)

  let rank = 0
  let prevCount = null
  let skip = 0

  const result = sortedEmojiCounts.map(([emoji, count]) => {
    if (count !== prevCount) {
      rank += skip + 1
      skip = 0
    } else {
      skip++
    }
    prevCount = count
    return `第${rank}位 ${emoji}: ${count}回`
  }).join('\n')

  return result
}


/**
 * 絵文字のランキングをSlackの指定されたチャンネルに投稿する関数。
 */
function postEmojiRankingToSlack() {
  const top10Emojis = collectEmojiRanking()
  const POST_CHANNEL_ID = PropertiesService.getScriptProperties().getProperty('POST_CHANNEL_ID')// 絵文字のランキングを投稿したいチャンネルのID

  const url = `https://slack.com/api/chat.postMessage`
  const options = {
    method: "post",
    contentType: "application/x-www-form-urlencoded",
    headers: { "Authorization": `Bearer ${SLACK_TOKEN}` },
    payload: {
      "channel": POST_CHANNEL_ID,
      "text": `直近1ヶ月 Slack　絵文字ランキング:\n${top10Emojis}`
    }
  };

  UrlFetchApp.fetch(url, options)
}


/**
 * 毎月1日の9時に絵文字のランキングをSlackに投稿するトリガーを設定する関数。
 */
function setMonthlyTrigger() {
  ScriptApp.newTrigger('postEmojiRankingToSlack')
    .timeBased()
    .onMonthDay(1)
    .atHour(9)
    .create()
}

