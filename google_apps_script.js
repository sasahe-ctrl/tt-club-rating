// ============================================================
//  TT CLUB & 广粤会 · 教练评分系统 · 后端逻辑
//  使用方式：把这整个文件的内容粘贴到 Google Apps Script 编辑器
//  必须绑定到一张 Google 表格（在表格里点"扩展程序→Apps Script"进入）
// ============================================================

const SHEET_NAME = 'ratings'; // 数据存在哪个工作表，不用改

// ------------------------------------------------------------
// 获取数据表；不存在时自动新建并写好表头
// ------------------------------------------------------------
function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow([
      '提交时间', 'UUID', '教练名', '课程名',
      '课程满意度', '教练专业度', '课堂氛围', '强度匹配', '推荐意愿',
      '文字反馈'
    ]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// ------------------------------------------------------------
// 检查：同一 UUID 对同一教练，今天（Asia/Shanghai 自然日）是否已提交过
// 返回 true = 已评过，false = 可以评
// ------------------------------------------------------------
function checkDuplicate(uuid, coach) {
  const sheet    = getSheet();
  const data     = sheet.getDataRange().getValues();
  const todayStr = Utilities.formatDate(new Date(), 'Asia/Shanghai', 'yyyy-MM-dd');

  for (let i = 1; i < data.length; i++) {
    const rowTime    = new Date(data[i][0]);
    const rowUUID    = data[i][1];
    const rowCoach   = data[i][2];
    const rowDateStr = Utilities.formatDate(rowTime, 'Asia/Shanghai', 'yyyy-MM-dd');

    if (rowUUID === uuid && rowCoach === coach && rowDateStr === todayStr) {
      return true;
    }
  }
  return false;
}

// ------------------------------------------------------------
// 统计：某教练本月收到多少个不重复 UUID 的评价
// 这就是成功页"你是本月第 N 位"的 N
// ------------------------------------------------------------
function countUniqueRaters(coach) {
  const sheet = getSheet();
  const data = sheet.getDataRange().getValues();
  const now = new Date();
  const firstOfMonth = new Date(now.getFullYear(), now.getMonth(), 1);
  const uuids = new Set();

  for (let i = 1; i < data.length; i++) {
    const rowTime  = new Date(data[i][0]);
    const rowUUID  = data[i][1];
    const rowCoach = data[i][2];

    if (rowCoach === coach && rowTime >= firstOfMonth) {
      uuids.add(rowUUID);
    }
  }
  return uuids.size;
}

// ------------------------------------------------------------
// GET 请求 — 页面加载时调用，检查这个 UUID 能不能评分
// 调用方式：GAS网址?uuid=xxx&coach=xxx
// ------------------------------------------------------------
function doGet(e) {
  try {
    const uuid  = (e.parameter.uuid  || '').trim();
    const coach = (e.parameter.coach || '').trim();

    if (!uuid || !coach) {
      return makeResponse({ ok: false, error: 'missing_params' });
    }

    const blocked = checkDuplicate(uuid, coach);
    return makeResponse({ ok: true, canRate: !blocked });

  } catch (err) {
    return makeResponse({ ok: false, error: 'server_error' });
  }
}

// ------------------------------------------------------------
// POST 请求 — 会员点"提交"时调用，写入数据并返回本月计数
// Body 格式：text/plain 包着 JSON 字符串
// ------------------------------------------------------------
function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    const { uuid, coach, course, scores, feedback } = payload;

    // 最终防重校验（防止极端情况下绕过页面加载时的检查）
    if (checkDuplicate(uuid, coach)) {
      return makeResponse({ ok: false, error: 'duplicate' });
    }

    // 写入一行数据
    const sheet   = getSheet();
    const timeStr = Utilities.formatDate(
      new Date(), 'Asia/Shanghai', 'yyyy-MM-dd HH:mm:ss'
    );

    sheet.appendRow([
      timeStr,
      uuid,
      coach,
      course,
      scores.satisfaction,
      scores.professionalism,
      scores.atmosphere,
      scores.intensity,
      scores.recommendation,
      feedback || ''
    ]);

    // 统计本月不重复评价人数（已包含刚写入的这条）
    const count = countUniqueRaters(coach);
    return makeResponse({ ok: true, count });

  } catch (err) {
    return makeResponse({ ok: false, error: 'server_error' });
  }
}

// ------------------------------------------------------------
// 工具：统一包装成 JSON 格式返回
// ------------------------------------------------------------
function makeResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
