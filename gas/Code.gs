// ============================================================
// webtv インサイト自動集計 - Google Apps Script v3
// ============================================================
//
// 【概要】
// Instagram Graph API (v21+) 2025年仕様対応版
// 2025年1月〜4月のAPI廃止・置換に完全対応
//
// 【シート構成（7シート）】
// 1. 通常投稿       - 静止画・カルーセルのインサイト
// 2. リール         - Reelsのインサイト
// 3. ストーリーズ   - Storiesのインサイト
// 4. アカウント全体 - アカウントレベルの日次指標
// 5. 設定           - APIトークン等の設定値
// 6. ログ           - 実行ログ
// 7. タイトル       - ダッシュボードからのカスタムタイトル
//
// 【2025年API変更対応メモ】
// 投稿: impressions→views, engagement→total_interactions, plays→views
// ストーリー: impressions→views, exits/taps→navigation
// アカウント: profile_views/website_clicks廃止 → 新指標群へ移行
//
// 【トリガー設定】
// - collectDailyAll()        → 毎日 7:00
// - refreshAccessToken()     → 毎月1日
// ============================================================


// ── 定数 ──────────────────────────────────

const SHEET_NAMES = {
  FEED: '通常投稿',
  REELS: 'リール',
  STORIES: 'ストーリーズ',
  ACCOUNT: 'アカウント全体',
  SETTINGS: '設定',
  LOG: 'ログ',
  TITLES: 'タイトル',
};

const API_VERSION = 'v21.0';
const API_BASE = `https://graph.facebook.com/${API_VERSION}`;

// 各シートのヘッダー（2025年API仕様準拠）
const HEADERS = {
  FEED: [
    '投稿日', '内容', 'メディアタイプ',
    '閲覧数', 'リーチ', 'いいね', 'コメント', '保存数', 'シェア',
    'インタラクション合計', 'フォロー数', 'プロフィールアクティビティ', 'プロフィール訪問数',
    'メディアID', '最終更新',
  ],
  REELS: [
    '投稿日', '内容',
    '閲覧数', 'リーチ', 'いいね', 'コメント', '保存数', 'シェア',
    'インタラクション合計',
    '平均視聴時間(秒)', '総再生時間(秒)',
    'メディアID', '最終更新',
  ],
  STORIES: [
    '投稿日', '内容',
    '閲覧数', 'リーチ', 'シェア', 'インタラクション合計',
    'フォロー数', 'プロフィールアクティビティ', 'プロフィール訪問数', 'ナビゲーション',
    'メディアID', '最終更新',
  ],
  ACCOUNT: [
    '日付',
    'フォロワー数', 'フォロワー増減',
    '閲覧数', 'アクションを実行したアカウント',
    'インタラクション数', 'いいね数', 'コメント数', '保存数', 'シェア数',
    'プロフィールリンクタップ', 'フォロー/アンフォロー',
    'フォロワー都市TOP5', 'フォロワー国TOP5', 'フォロワー性別年齢',
  ],
};


// ── 設定の読み書き ─────────────────────────

function getSetting(key) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.SETTINGS);
  if (!sheet) {
    throw new Error('「設定」シートが見つかりません。先にメニューの「📊 webtv → 初期セットアップ」を実行してください。');
  }
  const data = sheet.getDataRange().getValues();
  for (const row of data) {
    if (row[0] === key) return row[1];
  }
  return null;
}

function setSetting(key, value) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.SETTINGS);
  if (!sheet) {
    throw new Error('「設定」シートが見つかりません。先に initialSetup() を実行してください。');
  }
  const data = sheet.getDataRange().getValues();
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === key) {
      sheet.getRange(i + 1, 2).setValue(value);
      return;
    }
  }
  sheet.appendRow([key, value]);
}


// ── ログ ──────────────────────────────────

function writeLog(level, message) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.LOG);
  if (!sheet) {
    Logger.log(`[${level}] ${message} (ログシート未作成)`);
    return;
  }
  sheet.appendRow([new Date(), level, message]);
  if (sheet.getLastRow() > 500) {
    sheet.deleteRows(2, sheet.getLastRow() - 300);
  }
  Logger.log(`[${level}] ${message}`);
}


// ── Instagram Graph API ──────────────────

function callApi(endpoint, params = {}) {
  const token = getSetting('ACCESS_TOKEN');
  if (!token) {
    throw new Error('ACCESS_TOKENが設定されていません。設定シートを確認してください。');
  }

  const queryParts = [`access_token=${encodeURIComponent(token)}`];
  for (const key in params) {
    queryParts.push(`${encodeURIComponent(key)}=${encodeURIComponent(params[key])}`);
  }
  const url = `${API_BASE}/${endpoint}?${queryParts.join('&')}`;

  const maxRetries = 3;
  for (let attempt = 1; attempt <= maxRetries; attempt++) {
    try {
      const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
      const code = response.getResponseCode();
      const body = JSON.parse(response.getContentText());

      if (code === 200) return body;

      if (code === 401 || (body.error && body.error.code === 190)) {
        writeLog('ERROR', `アクセストークン期限切れ: ${body.error?.message}`);
        throw new Error('アクセストークンが期限切れです。refreshAccessToken()を実行してください。');
      }

      if (code === 429) {
        writeLog('WARN', `レート制限（${attempt}/${maxRetries}回目）。60秒待機...`);
        Utilities.sleep(60000);
        continue;
      }

      writeLog('ERROR', `API Error ${code}: ${JSON.stringify(body.error)}`);
      if (attempt === maxRetries) throw new Error(`API Error: ${body.error?.message}`);
      Utilities.sleep(5000 * attempt);

    } catch (e) {
      if (e.message.includes('アクセストークン')) throw e;
      if (attempt === maxRetries) throw e;
      writeLog('WARN', `リトライ ${attempt}/${maxRetries}: ${e.message}`);
      Utilities.sleep(5000 * attempt);
    }
  }
}

/**
 * 個別メトリクスを安全に取得（非対応メトリクスでも他に影響しない）
 */
function safeGetInsight(mediaId, metric) {
  try {
    const response = callApi(`${mediaId}/insights`, { metric: metric });
    if (response.data && response.data.length > 0) {
      return response.data[0].values[0].value;
    }
  } catch (e) {
    // 非対応メトリクスの場合は静かにスキップ
  }
  return 0;
}

/**
 * 複数メトリクスを一括取得
 */
function safeGetInsights(mediaId, metrics) {
  const result = {};
  try {
    const response = callApi(`${mediaId}/insights`, { metric: metrics });
    for (const item of response.data) {
      result[item.name] = item.values[0].value;
    }
  } catch (e) {
    writeLog('WARN', `インサイト一括取得失敗 (${mediaId}, ${metrics}): ${e.message}`);
  }
  return result;
}


// ── ヘルパー ─────────────────────────────

function getExistingMediaIds(sheet, mediaIdColIndex) {
  const ids = new Set();
  if (sheet.getLastRow() < 2) return ids;
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][mediaIdColIndex]) ids.add(String(data[i][mediaIdColIndex]));
  }
  return ids;
}

function updateRow(sheet, mediaId, mediaIdColIndex, updates) {
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][mediaIdColIndex]) === String(mediaId)) {
      const row = i + 1;
      for (const [col, val] of Object.entries(updates)) {
        sheet.getRange(row, parseInt(col)).setValue(val);
      }
      return true;
    }
  }
  return false;
}


// ══════════════════════════════════════════
// メイン処理①：通常投稿（静止画・カルーセル）
// ──────────────────────────────────────────
// 使用メトリクス:
//   views, reach, saved, shares,
//   total_interactions, follows, profile_activity, profile_visits
// ══════════════════════════════════════════

function collectFeedInsights() {
  const igUserId = getSetting('IG_USER_ID');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.FEED);

  if (sheet.getLastRow() === 0) sheet.appendRow(HEADERS.FEED);

  const existingIds = getExistingMediaIds(sheet, 13); // メディアID = N列(index 13)

  const mediaResponse = callApi(`${igUserId}/media`, {
    fields: 'id,caption,timestamp,media_type,media_product_type,like_count,comments_count',
    limit: '100',
  });

  let newCount = 0, updateCount = 0;

  for (const media of mediaResponse.data) {
    if (media.media_product_type === 'REELS' || media.media_product_type === 'STORY') continue;

    const mediaType = media.media_type === 'CAROUSEL_ALBUM' ? 'カルーセル' : '静止画';

    // 確実に取れるメトリクスを一括取得
    const base = safeGetInsights(media.id, 'reach,saved');
    // 個別取得
    const views = safeGetInsight(media.id, 'views');
    const shares = safeGetInsight(media.id, 'shares');
    const totalInteractions = safeGetInsight(media.id, 'total_interactions');
    const follows = safeGetInsight(media.id, 'follows');
    const profileActivity = safeGetInsight(media.id, 'profile_activity');
    const profileVisits = safeGetInsight(media.id, 'profile_visits');

    const postDate = Utilities.formatDate(new Date(media.timestamp), 'Asia/Tokyo', 'yyyy/M/d');
    const caption = (media.caption || '').substring(0, 100);

    if (existingIds.has(media.id)) {
      const daysSince = (new Date() - new Date(media.timestamp)) / 86400000;
      if (daysSince <= 30) {
        updateRow(sheet, media.id, 13, {
          4: views,                         // 閲覧数
          5: base.reach || 0,               // リーチ
          6: media.like_count || 0,         // いいね
          7: media.comments_count || 0,     // コメント
          8: base.saved || 0,               // 保存数
          9: shares,                        // シェア
          10: totalInteractions,            // インタラクション合計
          11: follows,                      // フォロー数
          12: profileActivity,              // プロフィールアクティビティ
          13: profileVisits,                // プロフィール訪問数
          15: new Date(),                   // 最終更新
        });
        updateCount++;
      }
    } else {
      sheet.appendRow([
        postDate, caption, mediaType,
        views, base.reach || 0, media.like_count || 0, media.comments_count || 0,
        base.saved || 0, shares,
        totalInteractions, follows, profileActivity, profileVisits,
        media.id, new Date(),
      ]);
      newCount++;
    }
    Utilities.sleep(1000);
  }

  writeLog('INFO', `通常投稿: 新規${newCount}件, 更新${updateCount}件`);
}


// ══════════════════════════════════════════
// メイン処理②：リール
// ──────────────────────────────────────────
// 使用メトリクス:
//   views (旧plays), reach, saved,
//   shares, total_interactions,
//   ig_reels_avg_watch_time, ig_reels_video_view_total_time
// ※ follows はリールでは非対応（通常投稿・ストーリーズのみ）
// ══════════════════════════════════════════

function collectReelsInsights() {
  const igUserId = getSetting('IG_USER_ID');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.REELS);

  if (sheet.getLastRow() === 0) sheet.appendRow(HEADERS.REELS);

  const existingIds = getExistingMediaIds(sheet, 11); // メディアID = L列(index 11)

  const mediaResponse = callApi(`${igUserId}/media`, {
    fields: 'id,caption,timestamp,media_type,media_product_type,like_count,comments_count',
    limit: '100',
  });

  let newCount = 0, updateCount = 0;

  for (const media of mediaResponse.data) {
    if (media.media_product_type !== 'REELS') continue;

    // 確実に取れるメトリクスを一括取得
    const base = safeGetInsights(media.id, 'reach,saved');
    // 個別取得
    const views = safeGetInsight(media.id, 'views');
    const shares = safeGetInsight(media.id, 'shares');
    const totalInteractions = safeGetInsight(media.id, 'total_interactions');
    const avgWatchTime = safeGetInsight(media.id, 'ig_reels_avg_watch_time');
    const totalWatchTime = safeGetInsight(media.id, 'ig_reels_video_view_total_time');

    const postDate = Utilities.formatDate(new Date(media.timestamp), 'Asia/Tokyo', 'yyyy/M/d');
    const caption = (media.caption || '').substring(0, 100);

    if (existingIds.has(media.id)) {
      const daysSince = (new Date() - new Date(media.timestamp)) / 86400000;
      if (daysSince <= 30) {
        updateRow(sheet, media.id, 11, {
          3: views,                         // 閲覧数
          4: base.reach || 0,               // リーチ
          5: media.like_count || 0,         // いいね
          6: media.comments_count || 0,     // コメント
          7: base.saved || 0,               // 保存数
          8: shares,                        // シェア
          9: totalInteractions,             // インタラクション合計
          10: avgWatchTime,                 // 平均視聴時間
          11: totalWatchTime,               // 総再生時間
          13: new Date(),                   // 最終更新
        });
        updateCount++;
      }
    } else {
      sheet.appendRow([
        postDate, caption,
        views, base.reach || 0, media.like_count || 0, media.comments_count || 0,
        base.saved || 0, shares,
        totalInteractions,
        avgWatchTime, totalWatchTime,
        media.id, new Date(),
      ]);
      newCount++;
    }
    Utilities.sleep(1000);
  }

  writeLog('INFO', `リール: 新規${newCount}件, 更新${updateCount}件`);
}


// ══════════════════════════════════════════
// メイン処理③：ストーリーズ
// ──────────────────────────────────────────
// 使用メトリクス:
//   views, reach, shares, total_interactions,
//   follows, profile_activity, profile_visits, navigation
// ══════════════════════════════════════════

function collectStoriesInsights() {
  const igUserId = getSetting('IG_USER_ID');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.STORIES);

  if (sheet.getLastRow() === 0) sheet.appendRow(HEADERS.STORIES);

  const existingIds = getExistingMediaIds(sheet, 10); // メディアID = K列(index 10)

  let storiesResponse;
  try {
    storiesResponse = callApi(`${igUserId}/stories`, {
      fields: 'id,caption,timestamp,media_type',
    });
  } catch (e) {
    writeLog('INFO', `ストーリーズ: 取得対象なし（現在公開中のストーリーズがない可能性）`);
    return;
  }

  if (!storiesResponse.data || storiesResponse.data.length === 0) {
    writeLog('INFO', 'ストーリーズ: 公開中のストーリーズなし');
    return;
  }

  let newCount = 0;

  for (const story of storiesResponse.data) {
    if (existingIds.has(story.id)) continue;

    const views = safeGetInsight(story.id, 'views');
    const reach = safeGetInsight(story.id, 'reach');
    const shares = safeGetInsight(story.id, 'shares');
    const totalInteractions = safeGetInsight(story.id, 'total_interactions');
    const follows = safeGetInsight(story.id, 'follows');
    const profileActivity = safeGetInsight(story.id, 'profile_activity');
    const profileVisits = safeGetInsight(story.id, 'profile_visits');
    const navigation = safeGetInsight(story.id, 'navigation');

    const postDate = Utilities.formatDate(new Date(story.timestamp), 'Asia/Tokyo', 'yyyy/M/d');
    const caption = (story.caption || '').substring(0, 100);

    sheet.appendRow([
      postDate, caption,
      views, reach, shares, totalInteractions,
      follows, profileActivity, profileVisits, navigation,
      story.id, new Date(),
    ]);
    newCount++;
    Utilities.sleep(1000);
  }

  writeLog('INFO', `ストーリーズ: 新規${newCount}件`);
}


// ══════════════════════════════════════════
// メイン処理④：アカウント全体の指標
// ──────────────────────────────────────────
// 使用メトリクス（2025年新指標群）:
//   views, accounts_engaged, total_interactions,
//   likes, comments, saves, shares,
//   profile_links_taps, follows_and_unfollows, follower_count
// ※ 旧: profile_views, website_clicks → 廃止済み
// ══════════════════════════════════════════

function collectAccountInsights() {
  const igUserId = getSetting('IG_USER_ID');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.ACCOUNT);

  if (sheet.getLastRow() === 0) sheet.appendRow(HEADERS.ACCOUNT);

  const today = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/M/d');

  // 同日の重複チェック
  if (sheet.getLastRow() >= 2) {
    const lastDate = sheet.getRange(sheet.getLastRow(), 1).getDisplayValue();
    if (lastDate === today) {
      writeLog('INFO', `アカウント全体: 本日分は記録済み（${today}）。スキップ`);
      return;
    }
  }

  // フォロワー数（プロフィール情報から取得）
  const profile = callApi(igUserId, { fields: 'followers_count' });
  const followersCount = profile.followers_count;

  let delta = 0;
  if (sheet.getLastRow() >= 2) {
    const prevCount = sheet.getRange(sheet.getLastRow(), 2).getValue();
    delta = followersCount - prevCount;
  }

  // アカウントレベル新指標（day period）
  let views = 0;
  let accountsEngaged = 0;
  let totalInteractions = 0;
  let likes = 0;
  let comments = 0;
  let saves = 0;
  let shares = 0;
  let profileLinksTaps = 0;
  let followsAndUnfollows = 0;

  // 一括取得できるメトリクスをまとめて取得
  try {
    const dayInsights = callApi(`${igUserId}/insights`, {
      metric: 'views,accounts_engaged,total_interactions,likes,comments,saves,shares,profile_links_taps,follows_and_unfollows',
      metric_type: 'total_value',
      period: 'day',
    });
    for (const item of dayInsights.data) {
      const val = item.total_value ? item.total_value.value : (item.values && item.values.length > 0 ? item.values[item.values.length - 1].value : 0);
      switch (item.name) {
        case 'views': views = val; break;
        case 'accounts_engaged': accountsEngaged = val; break;
        case 'total_interactions': totalInteractions = val; break;
        case 'likes': likes = val; break;
        case 'comments': comments = val; break;
        case 'saves': saves = val; break;
        case 'shares': shares = val; break;
        case 'profile_links_taps': profileLinksTaps = val; break;
        case 'follows_and_unfollows': followsAndUnfollows = val; break;
      }
    }
  } catch (e) {
    writeLog('WARN', `アカウント日次指標取得失敗: ${e.message}`);
    // 個別に取得を試みる
    const metrics = ['views', 'accounts_engaged', 'total_interactions', 'likes', 'comments', 'saves', 'shares', 'profile_links_taps', 'follows_and_unfollows'];
    const values = { views: 0, accounts_engaged: 0, total_interactions: 0, likes: 0, comments: 0, saves: 0, shares: 0, profile_links_taps: 0, follows_and_unfollows: 0 };
    for (const m of metrics) {
      try {
        const res = callApi(`${igUserId}/insights`, { metric: m, metric_type: 'total_value', period: 'day' });
        if (res.data && res.data.length > 0) {
          values[m] = res.data[0].total_value ? res.data[0].total_value.value : res.data[0].values[res.data[0].values.length - 1].value;
        }
      } catch (e2) { /* スキップ */ }
    }
    views = values.views;
    accountsEngaged = values.accounts_engaged;
    totalInteractions = values.total_interactions;
    likes = values.likes;
    comments = values.comments;
    saves = values.saves;
    shares = values.shares;
    profileLinksTaps = values.profile_links_taps;
    followsAndUnfollows = values.follows_and_unfollows;
  }

  // フォロワー属性（lifetime）- breakdownパラメータが必須
  let audienceCity = '';
  let audienceCountry = '';
  let audienceGenderAge = '';

  // 都市別
  try {
    const cityRes = callApi(`${igUserId}/insights`, {
      metric: 'follower_demographics',
      period: 'lifetime',
      metric_type: 'total_value',
      breakdown: 'city',
    });
    if (cityRes.data && cityRes.data.length > 0) {
      const results = cityRes.data[0].total_value.breakdowns[0].results;
      const sorted = results.sort((a, b) => b.value - a.value).slice(0, 5);
      audienceCity = sorted.map(r => `${r.dimension_values[0]}:${r.value}`).join(', ');
    }
  } catch (e) { /* スキップ */ }

  // 国別
  try {
    const countryRes = callApi(`${igUserId}/insights`, {
      metric: 'follower_demographics',
      period: 'lifetime',
      metric_type: 'total_value',
      breakdown: 'country',
    });
    if (countryRes.data && countryRes.data.length > 0) {
      const results = countryRes.data[0].total_value.breakdowns[0].results;
      const sorted = results.sort((a, b) => b.value - a.value).slice(0, 5);
      audienceCountry = sorted.map(r => `${r.dimension_values[0]}:${r.value}`).join(', ');
    }
  } catch (e) { /* スキップ */ }

  // 性別×年齢
  try {
    const genderRes = callApi(`${igUserId}/insights`, {
      metric: 'follower_demographics',
      period: 'lifetime',
      metric_type: 'total_value',
      breakdown: 'gender,age',
    });
    if (genderRes.data && genderRes.data.length > 0) {
      const results = genderRes.data[0].total_value.breakdowns[0].results;
      const sorted = results.sort((a, b) => b.value - a.value).slice(0, 8);
      audienceGenderAge = sorted.map(r => `${r.dimension_values.join('/')}:${r.value}`).join(', ');
    }
  } catch (e) { /* スキップ */ }

  sheet.appendRow([
    today,
    followersCount, delta,
    views, accountsEngaged,
    totalInteractions, likes, comments, saves, shares,
    profileLinksTaps, followsAndUnfollows,
    audienceCity, audienceCountry, audienceGenderAge,
  ]);

  writeLog('INFO', `アカウント全体: ${today} | ${followersCount}フォロワー (${delta >= 0 ? '+' : ''}${delta}) | 閲覧:${views} | インタラクション:${totalInteractions} | リンクタップ:${profileLinksTaps}`);
}


// ══════════════════════════════════════════
// 一括実行（日次トリガー用）
// ══════════════════════════════════════════

function collectDailyAll() {
  writeLog('INFO', '=== 日次一括収集 開始 ===');

  try { collectAccountInsights(); } catch (e) { writeLog('ERROR', `アカウント全体: ${e.message}`); }
  try { collectFeedInsights(); } catch (e) { writeLog('ERROR', `通常投稿: ${e.message}`); }
  try { collectReelsInsights(); } catch (e) { writeLog('ERROR', `リール: ${e.message}`); }
  try { collectStoriesInsights(); } catch (e) { writeLog('ERROR', `ストーリーズ: ${e.message}`); }

  writeLog('INFO', '=== 日次一括収集 完了 ===');
}


// ══════════════════════════════════════════
// アクセストークン更新
// ══════════════════════════════════════════

function refreshAccessToken() {
  try {
    const currentToken = getSetting('ACCESS_TOKEN');
    const appId = getSetting('APP_ID');
    const appSecret = getSetting('APP_SECRET');

    if (!currentToken || !appId || !appSecret) {
      writeLog('ERROR', 'トークン更新に必要な設定が不足しています');
      return;
    }

    const url = `https://graph.facebook.com/${API_VERSION}/oauth/access_token`
      + `?grant_type=fb_exchange_token`
      + `&client_id=${appId}`
      + `&client_secret=${appSecret}`
      + `&fb_exchange_token=${currentToken}`;

    const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    const body = JSON.parse(response.getContentText());

    if (body.access_token) {
      setSetting('ACCESS_TOKEN', body.access_token);
      setSetting('TOKEN_UPDATED', new Date().toISOString());
      writeLog('INFO', `アクセストークン更新成功。有効期限: ${body.expires_in}秒`);
    } else {
      writeLog('ERROR', `トークン更新失敗: ${JSON.stringify(body.error)}`);
      notifyTokenError(body.error?.message || '不明なエラー');
    }
  } catch (e) {
    writeLog('ERROR', `refreshAccessToken失敗: ${e.message}`);
    notifyTokenError(e.message);
  }
}

function notifyTokenError(errorMessage) {
  const email = getSetting('NOTIFY_EMAIL');
  if (!email) return;
  MailApp.sendEmail({
    to: email,
    subject: '【webtv】アクセストークンの更新に失敗しました',
    body: `webtvインサイト自動集計のアクセストークン更新に失敗しました。\n\n`
      + `エラー内容: ${errorMessage}\n\n`
      + `対応: Meta Business Suiteでトークンを再発行し、設定シートのACCESS_TOKENを更新してください。\n\n`
      + `スプレッドシート: ${SpreadsheetApp.getActiveSpreadsheet().getUrl()}`,
  });
}


// ══════════════════════════════════════════
// 初期セットアップ
// ══════════════════════════════════════════

function initialSetup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  for (const name of Object.values(SHEET_NAMES)) {
    if (!ss.getSheetByName(name)) {
      ss.insertSheet(name);
    }
  }

  const settingsSheet = ss.getSheetByName(SHEET_NAMES.SETTINGS);
  if (settingsSheet.getLastRow() === 0) {
    const settings = [
      ['設定キー', '値（ここを編集）'],
      ['IG_USER_ID', '（InstagramビジネスアカウントのユーザーID）'],
      ['ACCESS_TOKEN', '（Meta Business Suiteで取得した長期アクセストークン）'],
      ['APP_ID', '（MetaアプリのApp ID）'],
      ['APP_SECRET', '（MetaアプリのApp Secret）'],
      ['NOTIFY_EMAIL', '（トークンエラー通知先メールアドレス）'],
      ['TOKEN_UPDATED', ''],
    ];
    settingsSheet.getRange(1, 1, settings.length, 2).setValues(settings);
    settingsSheet.setColumnWidth(1, 200);
    settingsSheet.setColumnWidth(2, 500);
  }

  const sheetHeaderMap = {
    [SHEET_NAMES.FEED]: HEADERS.FEED,
    [SHEET_NAMES.REELS]: HEADERS.REELS,
    [SHEET_NAMES.STORIES]: HEADERS.STORIES,
    [SHEET_NAMES.ACCOUNT]: HEADERS.ACCOUNT,
  };

  for (const [sheetName, headers] of Object.entries(sheetHeaderMap)) {
    const s = ss.getSheetByName(sheetName);
    if (s && s.getLastRow() === 0) {
      s.appendRow(headers);
    }
  }

  // タイトルシートの初期化
  const titlesSheet = ss.getSheetByName(SHEET_NAMES.TITLES);
  if (titlesSheet && titlesSheet.getLastRow() === 0) {
    titlesSheet.appendRow(['キー', 'タイトル']);
  }

  const logSheet = ss.getSheetByName(SHEET_NAMES.LOG);
  if (logSheet.getLastRow() === 0) {
    logSheet.appendRow(['日時', 'レベル', 'メッセージ']);
  }

  writeLog('INFO', '初期セットアップ完了（v3: 2025年API仕様対応）');
  SpreadsheetApp.getUi().alert(
    'セットアップ完了（v3）',
    'シート構成:\n'
    + '・通常投稿（静止画・カルーセル）\n'
    + '・リール\n'
    + '・ストーリーズ\n'
    + '・アカウント全体\n'
    + '・タイトル（ダッシュボード用）\n\n'
    + '2025年API仕様に対応済みです。\n'
    + '「設定」シートにAPI情報を入力してください。',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}


// ══════════════════════════════════════════
// 接続テスト・デバッグ
// ══════════════════════════════════════════

function testConnection() {
  try {
    const igUserId = getSetting('IG_USER_ID');
    const data = callApi(igUserId, {
      fields: 'username,followers_count,media_count',
    });

    SpreadsheetApp.getUi().alert('接続テスト',
      `接続成功！\n\nユーザー名: ${data.username}\nフォロワー数: ${data.followers_count.toLocaleString()}\n投稿数: ${data.media_count}`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    writeLog('INFO', `接続テスト成功: @${data.username} (${data.followers_count}フォロワー)`);
  } catch (e) {
    SpreadsheetApp.getUi().alert('接続エラー',
      `エラー: ${e.message}\n\n設定シートのIG_USER_IDとACCESS_TOKENを確認してください。`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

/**
 * デバッグ用：最新3投稿のメトリクス対応状況を確認
 */
function debugInsights() {
  const igUserId = getSetting('IG_USER_ID');
  const token = getSetting('ACCESS_TOKEN');

  const mediaUrl = `${API_BASE}/${igUserId}/media?fields=id,caption,media_type,media_product_type,like_count,timestamp&limit=3&access_token=${encodeURIComponent(token)}`;
  const mediaRes = UrlFetchApp.fetch(mediaUrl, { muteHttpExceptions: true });
  const mediaBody = JSON.parse(mediaRes.getContentText());

  Logger.log('=== 最新3投稿（v3: 2025年API仕様） ===');
  for (const m of mediaBody.data) {
    Logger.log(`\n${m.media_product_type} | ${m.id} | ${(m.caption || '').substring(0, 30)}`);

    // 2025年公式仕様準拠のメトリクス
    const metrics = m.media_product_type === 'REELS'
      ? ['views', 'reach', 'saved', 'shares', 'total_interactions', 'ig_reels_avg_watch_time', 'ig_reels_video_view_total_time']
      : ['views', 'reach', 'saved', 'shares', 'total_interactions', 'follows', 'profile_activity', 'profile_visits'];

    for (const metric of metrics) {
      const insightsUrl = `${API_BASE}/${m.id}/insights?metric=${metric}&access_token=${encodeURIComponent(token)}`;
      const res = UrlFetchApp.fetch(insightsUrl, { muteHttpExceptions: true });
      const code = res.getResponseCode();
      if (code === 200) {
        const data = JSON.parse(res.getContentText());
        if (data.data && data.data.length > 0 && data.data[0].values && data.data[0].values.length > 0) {
          Logger.log(`  ✅ ${metric}: ${data.data[0].values[0].value}`);
        } else {
          Logger.log(`  ⚠️ ${metric}: HTTP 200 だがデータが空`);
        }
      } else {
        const err = JSON.parse(res.getContentText());
        Logger.log(`  ❌ ${metric}: HTTP ${code} - ${err.error?.message || ''}`);
      }
    }
  }

  // アカウント指標もテスト
  Logger.log('\n=== アカウント全体指標テスト ===');
  const accountMetrics = ['views', 'accounts_engaged', 'total_interactions', 'likes', 'comments', 'saves', 'shares', 'profile_links_taps', 'follows_and_unfollows'];
  for (const m of accountMetrics) {
    const url = `${API_BASE}/${igUserId}/insights?metric=${m}&metric_type=total_value&period=day&access_token=${encodeURIComponent(token)}`;
    const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    const code = res.getResponseCode();
    if (code === 200) {
      const data = JSON.parse(res.getContentText());
      if (data.data && data.data.length > 0) {
        const item = data.data[0];
        let val = 0;
        if (item.total_value) {
          val = item.total_value.value;
        } else if (item.values && item.values.length > 0) {
          val = item.values[item.values.length - 1].value;
        }
        Logger.log(`  ✅ ${m}: ${val}`);
      } else {
        Logger.log(`  ⚠️ ${m}: HTTP 200 だがデータが空`);
      }
    } else {
      const err = JSON.parse(res.getContentText());
      Logger.log(`  ❌ ${m}: HTTP ${code} - ${err.error?.message || ''}`);
    }
  }

  SpreadsheetApp.getUi().alert('デバッグ完了',
    '「表示」→「ログ」に2025年API仕様でのメトリクス対応状況を出力しました。',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}


// ══════════════════════════════════════════
// JSON APIエンドポイント（フロントエンド連携用）
// ══════════════════════════════════════════

function doGet(e) {
  const type = e.parameter.type || 'summary';
  const from = e.parameter.from || '';
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let result;

  try {
    // タイトル保存（GETベース: CORSプリフライト回避のため）
    if (e.parameter.action === 'setTitle' && e.parameter.key) {
      result = saveTitleData(ss, e.parameter.key, e.parameter.title || '');
    } else {
      switch (type) {
        case 'all': result = getAllData(ss, from); break;
        case 'account': result = getSheetData(ss, SHEET_NAMES.ACCOUNT, from); break;
        case 'feed': result = getSheetData(ss, SHEET_NAMES.FEED, from); break;
        case 'reels': result = getSheetData(ss, SHEET_NAMES.REELS, from); break;
        case 'stories': result = getSheetData(ss, SHEET_NAMES.STORIES, from); break;
        case 'titles': result = getTitlesData(ss); break;
        case 'summary': result = getSummaryData(ss, from); break;
        default: result = { error: 'Unknown type. Use: all, account, feed, reels, stories, titles, summary' };
      }
    }
  } catch (err) {
    result = { error: err.message };
  }

  return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    if (body.action === 'setTitle') {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      let sheet = ss.getSheetByName(SHEET_NAMES.TITLES);
      if (!sheet) {
        sheet = ss.insertSheet(SHEET_NAMES.TITLES);
        sheet.appendRow(['キー', 'タイトル']);
      }
      const data = sheet.getDataRange().getValues();
      let found = false;
      for (let i = 1; i < data.length; i++) {
        if (data[i][0] === body.key) {
          sheet.getRange(i + 1, 2).setValue(body.title);
          found = true;
          break;
        }
      }
      if (!found) {
        sheet.appendRow([body.key, body.title]);
      }
      return ContentService.createTextOutput(JSON.stringify({ success: true }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    return ContentService.createTextOutput(JSON.stringify({ error: 'Unknown action' }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function getAllData(ss, from) {
  return {
    type: 'all',
    account: getSheetData(ss, SHEET_NAMES.ACCOUNT, from),
    feed: getSheetData(ss, SHEET_NAMES.FEED, from),
    reels: getSheetData(ss, SHEET_NAMES.REELS, from),
    stories: getSheetData(ss, SHEET_NAMES.STORIES, from),
    titles: getTitlesData(ss),
  };
}

function getSheetData(ss, sheetName, from) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() < 2) return { type: sheetName, count: 0, data: [] };
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const rows = [];
  for (let i = 1; i < data.length; i++) {
    const dateStr = data[i][0];
    if (from && dateStr < from) continue;
    const row = {};
    for (let j = 0; j < headers.length; j++) {
      row[headers[j]] = data[i][j];
    }
    rows.push(row);
  }
  return { type: sheetName, count: rows.length, data: rows };
}

function saveTitleData(ss, key, title) {
  let sheet = ss.getSheetByName(SHEET_NAMES.TITLES);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAMES.TITLES);
    sheet.appendRow(['キー', 'タイトル']);
  }
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === key) {
      sheet.getRange(i + 1, 2).setValue(title);
      return { success: true, updated: true };
    }
  }
  sheet.appendRow([key, title]);
  return { success: true, created: true };
}

function getTitlesData(ss) {
  const sheet = ss.getSheetByName(SHEET_NAMES.TITLES);
  if (!sheet || sheet.getLastRow() < 2) return { data: {} };
  const data = sheet.getDataRange().getValues();
  const titles = {};
  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) titles[data[i][0]] = data[i][1];
  }
  return { data: titles };
}

function getSummaryData(ss, from) {
  const account = getSheetData(ss, SHEET_NAMES.ACCOUNT, from);
  const feed = getSheetData(ss, SHEET_NAMES.FEED, from);
  const reels = getSheetData(ss, SHEET_NAMES.REELS, from);

  return {
    type: 'summary',
    account: {
      days: account.data.length,
      latestFollowers: account.data.length > 0 ? account.data[account.data.length - 1]['フォロワー数'] : 0,
      firstFollowers: account.data.length > 0 ? account.data[0]['フォロワー数'] - account.data[0]['フォロワー増減'] : 0,
    },
    feed: {
      count: feed.data.length,
      totalViews: feed.data.reduce((s, r) => s + (r['閲覧数'] || 0), 0),
      totalReach: feed.data.reduce((s, r) => s + (r['リーチ'] || 0), 0),
      totalLikes: feed.data.reduce((s, r) => s + (r['いいね'] || 0), 0),
      totalSaves: feed.data.reduce((s, r) => s + (r['保存数'] || 0), 0),
    },
    reels: {
      count: reels.data.length,
      totalViews: reels.data.reduce((s, r) => s + (r['閲覧数'] || 0), 0),
      totalReach: reels.data.reduce((s, r) => s + (r['リーチ'] || 0), 0),
      totalSaves: reels.data.reduce((s, r) => s + (r['保存数'] || 0), 0),
      totalShares: reels.data.reduce((s, r) => s + (r['シェア'] || 0), 0),
    },
  };
}


// ══════════════════════════════════════════
// カスタムメニュー
// ══════════════════════════════════════════

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📊 webtv')
    .addItem('初期セットアップ', 'initialSetup')
    .addItem('接続テスト', 'testConnection')
    .addItem('インサイト取得テスト', 'debugInsights')
    .addSeparator()
    .addItem('全データを今すぐ取得', 'collectDailyAll')
    .addSeparator()
    .addItem('通常投稿のみ取得', 'collectFeedInsights')
    .addItem('リールのみ取得', 'collectReelsInsights')
    .addItem('ストーリーズのみ取得', 'collectStoriesInsights')
    .addItem('アカウント全体のみ取得', 'collectAccountInsights')
    .addSeparator()
    .addItem('アクセストークンを更新', 'refreshAccessToken')
    .addToUi();
}
