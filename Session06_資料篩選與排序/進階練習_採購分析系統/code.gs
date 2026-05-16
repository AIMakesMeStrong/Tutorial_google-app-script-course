// ============================================================
// 進階練習：智慧採購分析與供應商評比
// 對應：Session 6（filter、sort、資料摘要）
// ============================================================

/**
 * 供應商綜合評比（篩選 + 排序 + 統計）
 */
function 供應商評比() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("採購紀錄");
    if (!sheet) { SpreadsheetApp.getUi().alert("❌ 請先初始化"); return; }

    var 資料 = sheet.getDataRange().getValues();
    var 標題 = 資料[0];
    var 紀錄 = [];
    for (var i = 1; i < 資料.length; i++) {
      var obj = {};
      for (var j = 0; j < 標題.length; j++) obj[標題[j]] = 資料[i][j];
      紀錄.push(obj);
    }

    // 依供應商分組統計
    var 供應商統計 = {};
    紀錄.forEach(function(r) {
      var s = r["供應商"];
      if (!供應商統計[s]) {
        供應商統計[s] = { 訂單數: 0, 總金額: 0, 準時交貨: 0, 退貨次數: 0, 品質分數: [] };
      }
      供應商統計[s].訂單數++;
      供應商統計[s].總金額 += r["金額"];
      if (r["交貨狀態"] === "準時") 供應商統計[s].準時交貨++;
      if (r["退貨"] === "是") 供應商統計[s].退貨次數++;
      供應商統計[s].品質分數.push(r["品質評分"]);
    });

    // 計算綜合評分
    var 排名 = [];
    for (var name in 供應商統計) {
      var s = 供應商統計[name];
      var 準時率 = (s.準時交貨 / s.訂單數 * 100);
      var 退貨率 = (s.退貨次數 / s.訂單數 * 100);
      var 平均品質 = s.品質分數.reduce(function(a, b) { return a + b; }, 0) / s.品質分數.length;

      // 綜合評分 = 品質(40%) + 準時率(30%) + (100-退貨率)(20%) + 價格競爭力(10%)
      var 綜合 = 平均品質 * 0.4 + 準時率 * 0.3 + (100 - 退貨率) * 0.2 + 50 * 0.1;

      排名.push({
        供應商: name, 訂單數: s.訂單數, 總金額: s.總金額,
        準時率: 準時率, 退貨率: 退貨率, 品質: 平均品質, 綜合: 綜合
      });
    }

    // 依綜合評分排序
    排名.sort(function(a, b) { return b.綜合 - a.綜合; });

    // 產生評比報表
    var 評比表 = ss.getSheetByName("供應商評比");
    if (評比表) 評比表.clear(); else 評比表 = ss.insertSheet("供應商評比");

    評比表.getRange("A1").setValue("🏆 供應商綜合評比").setFontSize(16).setFontWeight("bold");
    評比表.getRange("A2").setValue("評比日期：" + Utilities.formatDate(new Date(), "Asia/Taipei", "yyyy/MM/dd"));

    var 表頭 = [["排名", "供應商", "訂單數", "總金額", "準時率", "退貨率", "品質", "綜合評分", "等級"]];
    評比表.getRange(4, 1, 1, 9).setValues(表頭);
    評比表.getRange(4, 1, 1, 9).setBackground("#1565c0").setFontColor("#fff").setFontWeight("bold");

    排名.forEach(function(r, idx) {
      var 等級 = r.綜合 >= 80 ? "⭐ A" : r.綜合 >= 60 ? "🔵 B" : "🔴 C";
      評比表.getRange(5 + idx, 1, 1, 9).setValues([[
        idx + 1, r.供應商, r.訂單數, r.總金額,
        (r.準時率).toFixed(1) + "%", (r.退貨率).toFixed(1) + "%",
        r.品質.toFixed(1), r.綜合.toFixed(1), 等級
      ]]);
      if (idx === 0) 評比表.getRange(5 + idx, 1, 1, 9).setBackground("#e8f5e9"); // 冠軍綠
    });

    評比表.getRange(5, 4, 排名.length, 1).setNumberFormat("#,##0");
    for (var c = 1; c <= 9; c++) 評比表.autoResizeColumn(c);

    SpreadsheetApp.getUi().alert("✅ 供應商評比完成！冠軍：" + 排名[0].供應商);

  } catch (錯誤) { Logger.log("❌ " + 錯誤.message); }
}

/**
 * 採購需求預測（根據歷史消耗趨勢）
 */
function 採購預測() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("採購紀錄");
  if (!sheet) return;

  var 資料 = sheet.getDataRange().getValues();
  var 品項統計 = {};

  for (var i = 1; i < 資料.length; i++) {
    var 品項 = 資料[i][1]; // B: 品項
    var 數量 = 資料[i][3]; // D: 數量
    if (!品項統計[品項]) 品項統計[品項] = [];
    品項統計[品項].push(數量);
  }

  var 預測 = [];
  for (var name in 品項統計) {
    var 歷史 = 品項統計[name];
    var 平均 = 歷史.reduce(function(a, b) { return a + b; }, 0) / 歷史.length;
    var 建議量 = Math.ceil(平均 * 1.2); // 多 20% 安全庫存

    預測.push(name + "：月均 " + Math.round(平均) + " → 建議採購 " + 建議量);
  }

  SpreadsheetApp.getUi().alert("📊 採購預測\n\n" + 預測.join("\n"));
}

function 初始化採購資料() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("採購紀錄");
  if (!sheet) sheet = ss.insertSheet("採購紀錄"); else sheet.clear();

  var 標題 = [["供應商", "品項", "單價", "數量", "金額", "交貨狀態", "退貨", "品質評分", "日期"]];
  var 供應商 = ["宏達文具", "大同辦公", "金鼎耗材", "永豐科技", "佳能事務"];
  var 品項 = ["A4影印紙", "碳粉匣", "原子筆", "資料夾", "白板筆"];
  var 資料 = [];

  for (var i = 0; i < 40; i++) {
    var s = 供應商[Math.floor(Math.random() * 供應商.length)];
    var p = 品項[Math.floor(Math.random() * 品項.length)];
    var 單價 = [150, 1200, 15, 25, 45][品項.indexOf(p)];
    var 數量 = Math.floor(Math.random() * 50) + 5;
    var 交貨 = Math.random() > 0.2 ? "準時" : "延遲";
    var 退貨 = Math.random() > 0.85 ? "是" : "否";
    var 品質 = Math.floor(Math.random() * 40) + 60;
    var 日期 = new Date(2026, Math.floor(Math.random() * 4), Math.floor(Math.random() * 28) + 1);

    資料.push([s, p, 單價, 數量, 單價 * 數量, 交貨, 退貨, 品質, 日期]);
  }

  sheet.getRange(1, 1, 1, 9).setValues(標題);
  sheet.getRange(2, 1, 資料.length, 9).setValues(資料);
  sheet.getRange("A1:I1").setBackground("#6a1b9a").setFontColor("#fff").setFontWeight("bold");
  sheet.getRange("E2:E41").setNumberFormat("#,##0");
  sheet.getRange("I2:I41").setNumberFormat("yyyy/mm/dd");
  sheet.setFrozenRows(1);
  for (var c = 1; c <= 9; c++) sheet.autoResizeColumn(c);

  SpreadsheetApp.getUi().alert("✅ 40 筆採購紀錄已建立！");
}

// ============================================================
// 進階挑戰 1：價格趨勢分析（同品項不同月份的價格走勢）
// ============================================================

/**
 * 分析同品項在不同月份的單價走勢，標記漲幅異常的品項
 */
function 價格趨勢分析() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("採購紀錄");
    if (!sheet) { SpreadsheetApp.getUi().alert("❌ 請先初始化採購資料"); return; }

    var 資料 = sheet.getDataRange().getValues();
    var 標題 = 資料[0];

    // 依「品項 + 月份」分組，收集每筆單價
    var 趨勢 = {}; // { 品項: { 月份: [單價, 單價, ...] } }
    for (var i = 1; i < 資料.length; i++) {
      var 品項 = 資料[i][1]; // B欄：品項
      var 單價 = 資料[i][2]; // C欄：單價
      var 日期 = new Date(資料[i][8]); // I欄：日期
      var 月份 = 日期.getMonth() + 1;

      if (!趨勢[品項]) 趨勢[品項] = {};
      if (!趨勢[品項][月份]) 趨勢[品項][月份] = [];
      趨勢[品項][月份].push(單價);
    }

    // 計算每品項每月平均單價
    var 所有月份 = [1, 2, 3, 4];
    var 結果 = [];

    for (var 品項名 in 趨勢) {
      var row = [品項名];
      var 第一月均價 = null;
      var 最後月均價 = null;

      所有月份.forEach(function(m) {
        var prices = 趨勢[品項名][m];
        if (prices && prices.length > 0) {
          var avg = prices.reduce(function(a, b) { return a + b; }, 0) / prices.length;
          row.push(Math.round(avg));
          if (第一月均價 === null) 第一月均價 = avg;
          最後月均價 = avg;
        } else {
          row.push("-");
        }
      });

      // 計算整體漲跌幅
      var 漲跌幅 = (第一月均價 && 最後月均價) ?
        ((最後月均價 - 第一月均價) / 第一月均價 * 100).toFixed(1) + "%" : "N/A";
      row.push(漲跌幅);

      結果.push(row);
    }

    // 建立「價格趨勢」工作表
    var 趨勢表 = ss.getSheetByName("價格趨勢");
    if (趨勢表) 趨勢表.clear(); else 趨勢表 = ss.insertSheet("價格趨勢");

    趨勢表.getRange("A1").setValue("📈 品項價格趨勢分析").setFontSize(16).setFontWeight("bold");
    趨勢表.getRange("A2").setValue("分析日期：" + Utilities.formatDate(new Date(), "Asia/Taipei", "yyyy/MM/dd"));

    var 表頭 = [["品項", "1月均價", "2月均價", "3月均價", "4月均價", "漲跌幅"]];
    趨勢表.getRange(4, 1, 1, 6).setValues(表頭);
    趨勢表.getRange(4, 1, 1, 6).setBackground("#e65100").setFontColor("#fff").setFontWeight("bold");

    if (結果.length > 0) {
      趨勢表.getRange(5, 1, 結果.length, 6).setValues(結果);

      // 條件格式：漲幅超過 10% 標紅色警告
      結果.forEach(function(row, idx) {
        var 漲跌 = parseFloat(row[5]);
        if (!isNaN(漲跌) && 漲跌 > 10) {
          趨勢表.getRange(5 + idx, 1, 1, 6).setBackground("#ffebee").setFontColor("#c62828");
        } else if (!isNaN(漲跌) && 漲跌 < -10) {
          趨勢表.getRange(5 + idx, 1, 1, 6).setBackground("#e8f5e9").setFontColor("#2e7d32");
        }
      });
    }

    for (var c = 1; c <= 6; c++) 趨勢表.autoResizeColumn(c);
    SpreadsheetApp.getUi().alert("✅ 價格趨勢分析完成！請查看「價格趨勢」工作表");

  } catch (錯誤) { Logger.log("❌ " + 錯誤.message); SpreadsheetApp.getUi().alert("❌ " + 錯誤.message); }
}

// ============================================================
// 進階挑戰 2：自動產生採購建議書（含推薦供應商與預估金額）
// ============================================================

/**
 * 根據供應商評比 + 採購預測，自動生成下季採購建議書
 */
function 產生採購建議書() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("採購紀錄");
    if (!sheet) { SpreadsheetApp.getUi().alert("❌ 請先初始化採購資料"); return; }

    var 資料 = sheet.getDataRange().getValues();
    var 標題 = 資料[0];
    var 紀錄 = [];
    for (var i = 1; i < 資料.length; i++) {
      var obj = {};
      for (var j = 0; j < 標題.length; j++) obj[標題[j]] = 資料[i][j];
      紀錄.push(obj);
    }

    // --- 供應商評比邏輯（取得排名） ---
    var 供應商統計 = {};
    紀錄.forEach(function(r) {
      var s = r["供應商"];
      if (!供應商統計[s]) {
        供應商統計[s] = { 訂單數: 0, 總金額: 0, 準時交貨: 0, 退貨次數: 0, 品質分數: [] };
      }
      供應商統計[s].訂單數++;
      供應商統計[s].總金額 += r["金額"];
      if (r["交貨狀態"] === "準時") 供應商統計[s].準時交貨++;
      if (r["退貨"] === "是") 供應商統計[s].退貨次數++;
      供應商統計[s].品質分數.push(r["品質評分"]);
    });

    var 排名 = [];
    for (var name in 供應商統計) {
      var s = 供應商統計[name];
      var 準時率 = s.準時交貨 / s.訂單數 * 100;
      var 退貨率 = s.退貨次數 / s.訂單數 * 100;
      var 平均品質 = s.品質分數.reduce(function(a, b) { return a + b; }, 0) / s.品質分數.length;
      var 綜合 = 平均品質 * 0.4 + 準時率 * 0.3 + (100 - 退貨率) * 0.2 + 50 * 0.1;
      var 等級 = 綜合 >= 80 ? "A" : 綜合 >= 60 ? "B" : "C";
      排名.push({ 供應商: name, 綜合: 綜合, 等級: 等級 });
    }
    排名.sort(function(a, b) { return b.綜合 - a.綜合; });

    // --- 各品項的建議採購量與最近單價 ---
    var 品項統計 = {};
    紀錄.forEach(function(r) {
      var p = r["品項"];
      if (!品項統計[p]) 品項統計[p] = { 數量: [], 單價: [], 供應商: {} };
      品項統計[p].數量.push(r["數量"]);
      品項統計[p].單價.push(r["單價"]);
      // 記錄每個供應商對此品項的供貨次數
      if (!品項統計[p].供應商[r["供應商"]]) 品項統計[p].供應商[r["供應商"]] = 0;
      品項統計[p].供應商[r["供應商"]]++;
    });

    // 為每個品項匹配評分最高且有供貨紀錄的供應商
    var 建議列表 = [];
    var 總預估金額 = 0;

    for (var 品項名 in 品項統計) {
      var info = 品項統計[品項名];
      var 平均量 = info.數量.reduce(function(a, b) { return a + b; }, 0) / info.數量.length;
      var 建議量 = Math.ceil(平均量 * 1.2);
      var 最近單價 = info.單價[info.單價.length - 1];
      var 預估金額 = 建議量 * 最近單價;

      // 從排名中找出有供貨此品項的最佳供應商
      var 推薦供應商 = "N/A";
      var 推薦等級 = "-";
      for (var r = 0; r < 排名.length; r++) {
        if (info.供應商[排名[r].供應商]) {
          推薦供應商 = 排名[r].供應商;
          推薦等級 = 排名[r].等級;
          break;
        }
      }

      建議列表.push([品項名, 推薦供應商, 推薦等級, 建議量, 最近單價, 預估金額]);
      總預估金額 += 預估金額;
    }

    // --- 建立「採購建議書」工作表 ---
    var 建議表 = ss.getSheetByName("採購建議書");
    if (建議表) 建議表.clear(); else 建議表 = ss.insertSheet("採購建議書");

    // 標題區
    建議表.getRange("A1").setValue("📋 下季採購建議書").setFontSize(18).setFontWeight("bold");
    建議表.getRange("A2").setValue("產生日期：" + Utilities.formatDate(new Date(), "Asia/Taipei", "yyyy/MM/dd"));
    建議表.getRange("A3").setValue("建議採購期間：2026 Q2");

    // 推薦供應商摘要
    建議表.getRange("A5").setValue("🏆 推薦供應商排名").setFontSize(13).setFontWeight("bold");
    建議表.getRange(6, 1, 1, 3).setValues([["供應商", "綜合評分", "等級"]]);
    建議表.getRange(6, 1, 1, 3).setBackground("#1565c0").setFontColor("#fff").setFontWeight("bold");

    排名.forEach(function(r, idx) {
      建議表.getRange(7 + idx, 1, 1, 3).setValues([[r.供應商, r.綜合.toFixed(1), r.等級]]);
      if (r.等級 === "A") 建議表.getRange(7 + idx, 3).setFontColor("#2e7d32").setFontWeight("bold");
      if (r.等級 === "C") 建議表.getRange(7 + idx, 3).setFontColor("#c62828").setFontWeight("bold");
    });

    // 品項建議明細
    var 明細起始列 = 7 + 排名.length + 2;
    建議表.getRange(明細起始列 - 1, 1).setValue("📦 品項採購建議明細").setFontSize(13).setFontWeight("bold");
    建議表.getRange(明細起始列, 1, 1, 6).setValues([["品項", "推薦供應商", "等級", "建議數量", "參考單價", "預估金額"]]);
    建議表.getRange(明細起始列, 1, 1, 6).setBackground("#6a1b9a").setFontColor("#fff").setFontWeight("bold");

    建議列表.forEach(function(row, idx) {
      建議表.getRange(明細起始列 + 1 + idx, 1, 1, 6).setValues([row]);
    });

    // 總預估金額
    var 合計列 = 明細起始列 + 1 + 建議列表.length + 1;
    建議表.getRange(合計列, 5).setValue("總預估金額").setFontWeight("bold");
    建議表.getRange(合計列, 6).setValue(總預估金額).setNumberFormat("#,##0").setFontWeight("bold").setFontSize(14).setFontColor("#1565c0");

    // 格式化金額欄
    建議表.getRange(明細起始列 + 1, 5, 建議列表.length, 2).setNumberFormat("#,##0");
    for (var c = 1; c <= 6; c++) 建議表.autoResizeColumn(c);

    SpreadsheetApp.getUi().alert("✅ 採購建議書已產生！\n\n下季預估總金額：$" + 總預估金額.toLocaleString());

  } catch (錯誤) { Logger.log("❌ " + 錯誤.message); SpreadsheetApp.getUi().alert("❌ " + 錯誤.message); }
}

// ============================================================
// 進階挑戰 3：Gmail 自動發送評比結果給供應商
// ============================================================

/**
 * 將供應商評比結果以 HTML 格式透過 Gmail 寄出
 * 收件人由使用者透過提示框輸入
 */
function 發送評比郵件() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var 評比表 = ss.getSheetByName("供應商評比");
    if (!評比表) {
      SpreadsheetApp.getUi().alert("❌ 請先執行「供應商評比」產生報表");
      return;
    }

    // 讀取評比資料（第 4 列是表頭，第 5 列開始是資料）
    var 表頭 = 評比表.getRange(4, 1, 1, 9).getValues()[0];
    var lastRow = 評比表.getLastRow();
    if (lastRow < 5) {
      SpreadsheetApp.getUi().alert("❌ 評比表中沒有資料");
      return;
    }
    var 資料 = 評比表.getRange(5, 1, lastRow - 4, 9).getValues();

    // 請使用者輸入收件人 Email
    var ui = SpreadsheetApp.getUi();
    var response = ui.prompt(
      "📧 發送評比報告",
      "請輸入收件人 Email 地址：",
      ui.ButtonSet.OK_CANCEL
    );

    if (response.getSelectedButton() !== ui.Button.OK) return;
    var email = response.getResponseText().trim();
    if (!email || email.indexOf("@") === -1) {
      ui.alert("❌ 請輸入有效的 Email 地址");
      return;
    }

    // 組合 HTML 郵件
    var today = Utilities.formatDate(new Date(), "Asia/Taipei", "yyyy/MM/dd");
    var html = '<div style="font-family:Arial,sans-serif;max-width:600px;margin:0 auto">';
    html += '<h2 style="color:#1565c0">🏆 供應商綜合評比報告</h2>';
    html += '<p>評比日期：' + today + '</p>';
    html += '<table style="border-collapse:collapse;width:100%;font-size:13px">';
    html += '<tr style="background:#1565c0;color:#fff">';
    表頭.forEach(function(h) {
      html += '<th style="padding:8px;border:1px solid #ddd">' + h + '</th>';
    });
    html += '</tr>';

    資料.forEach(function(row, idx) {
      var bgColor = idx === 0 ? '#e8f5e9' : (idx % 2 === 0 ? '#f5f5f5' : '#fff');
      html += '<tr style="background:' + bgColor + '">';
      row.forEach(function(cell) {
        html += '<td style="padding:6px;border:1px solid #ddd;text-align:center">' + cell + '</td>';
      });
      html += '</tr>';
    });

    html += '</table>';
    html += '<p style="color:#666;margin-top:16px;font-size:12px">此報告由智慧採購分析系統自動產生</p>';
    html += '</div>';

    // 發送郵件
    MailApp.sendEmail({
      to: email,
      subject: "【供應商評比報告】" + today,
      htmlBody: html
    });

    ui.alert("✅ 評比報告已寄送至：" + email);

  } catch (錯誤) { Logger.log("❌ " + 錯誤.message); SpreadsheetApp.getUi().alert("❌ " + 錯誤.message); }
}

// ============================================================
// 自訂選單
// ============================================================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("🤖 智慧採購管理")
    .addItem("📦 初始化採購資料", "初始化採購資料")
    .addItem("🏆 供應商評比", "供應商評比")
    .addItem("📊 採購預測", "採購預測")
    .addSeparator()
    .addItem("📈 價格趨勢分析", "價格趨勢分析")
    .addItem("📋 產生採購建議書", "產生採購建議書")
    .addItem("📧 發送評比郵件", "發送評比郵件")
    .addToUi();
}
