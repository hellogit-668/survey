// ============================================
// 1. 在 Google Sheet 中点击: 扩展程序 → Apps Script
// 2. 粘贴以下全部代码，替换默认内容
// 3. 点击顶部"部署" → "新建部署" → 类型选"Web 应用"
// 4. 访问权限设为"任何人"，点击"部署"
// 5. 复制生成的 Web 应用 URL，填入 survey.html 的 SUBMIT_URL
// ============================================

function doPost(e) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var data = JSON.parse(e.postData.contents);
    var answers = data.answers;

    // 如果表头为空，先写入表头
    if (sheet.getLastRow() === 0) {
      sheet.appendRow([
        '提交时间',
        '1.是否购买过户外品牌产品',
        '2.性别',
        '3.年龄',
        '4.教育程度',
        '5.每月可支配消费金额',
        '6.购买主要渠道',
        '7.购买频率',
        '8.多渠道便捷获取产品信息',
        '9.线上线下价格服务统一',
        '10.客服响应及时',
        '11.自由切换线上线下渠道',
        '12.公示环保材质信息',
        '13.绿色包装回收公益',
        '14.普及环保知识理念',
        '15.绿色营销真实可信',
        '16.全渠道优质体验性价比',
        '17.满足绿色消费心理需求',
        '18.提升幸福感认同感',
        '19.综合良好消费价值',
        '20.优先选择该品牌',
        '21.即时购买决策',
        '22.增加购买频次金额',
        '23.持续复购',
        '24.主动推荐他人'
      ]);
    }

    var row = [new Date()];
    for (var i = 1; i <= 24; i++) {
      var val = answers[i];
      if (val === undefined || val === null) {
        row.push('');
      } else {
        row.push(val);
      }
    }
    sheet.appendRow(row);

    return ContentService.createTextOutput(JSON.stringify({ success: true }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet() {
  return ContentService.createTextOutput('Survey API OK');
}
