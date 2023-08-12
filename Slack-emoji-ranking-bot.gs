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
 * メッセージから絵文字を集計する関数。
 * @param {Object} message - Slackのメッセージオブジェクト。
 * @param {Object} emojiCounts - 絵文字の使用回数を集計するオブジェクト。
 */
function countEmojisFromMessage(message, emojiCounts) {
  // 会話のテキスト内の絵文字を集計
  const emojisInText = (message.text || "").match(/:(\w+):/g) || []
  emojisInText.forEach(emoji => {
    if (!emojiCounts[emoji]) {
      emojiCounts[emoji] = 0
    }
    emojiCounts[emoji]++
  });

  // リアクションとして追加された絵文字を集計
  if (message.reactions) {
    message.reactions.forEach(reaction => {
      const emoji = `:${reaction.name}:`
      if (!emojiCounts[emoji]) {
        emojiCounts[emoji] = 0
      }
      // リアクションの回数（count）を絵文字の使用回数に追加
      emojiCounts[emoji] += reaction.count
    });
  }
}

/**
 * 特定のチャンネルのメッセージ履歴を取得する関数。
 * @param {string} channelId - SlackのチャンネルID。
 * @param {number} oldestTimestamp - 集計開始のUNIXタイムスタンプ。
 * @return {Array} チャンネルのメッセージ履歴。
 */
function fetchChannelHistory(channelId, oldestTimestamp) {
  let hasMore = true
  let cursor = null
  const allMessages = []

  while (hasMore) {
    const url = `https://slack.com/api/conversations.history`
    const options = {
      method: "get",
      contentType: "application/x-www-form-urlencoded",
      headers: { "Authorization": `Bearer ${SLACK_TOKEN}` },
      payload: {
        "channel": channelId,
        "limit": 1000,
        "oldest": oldestTimestamp,
        "cursor": cursor
      }
    }

    const historyResponse = UrlFetchApp.fetch(url, options)
    const data = JSON.parse(historyResponse.getContentText())

    if (data.messages) {
      allMessages.push(...data.messages)
    }

    hasMore = data.has_more
    cursor = data.response_metadata ? data.response_metadata.next_cursor : null
  }

  return allMessages
}

/**
 * 絵文字のランキングを生成する関数。
 * @param {Object} emojiCounts - 絵文字の使用回数を集計するオブジェクト。
 * @return {string} トップ10の絵文字とその使用回数の文字列。
 */
function generateEmojiRanking(emojiCounts) {
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
 * Slackのワークスペースでの会話において、直近1ヶ月間の利用されている絵文字のランキングを集計する関数。
 * @return {string} トップ10の絵文字とその使用回数の文字列。
 */
function collectEmojiRanking() {
  const channelIds = getAllChannelIds()
  const emojiCounts = {}

  // 現在の日時から1ヶ月前のUNIXタイムスタンプを計算
  const today = new Date()
  today.setMonth(today.getMonth() - 1)
  const oldestTimestamp = Math.floor(today.getTime() / 1000)

  channelIds.forEach(channelId => {
    const messages = fetchChannelHistory(channelId, oldestTimestamp)
    messages.forEach(message => {
      countEmojisFromMessage(message, emojiCounts)
    })
  })

  return generateEmojiRanking(emojiCounts)
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
  }

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

