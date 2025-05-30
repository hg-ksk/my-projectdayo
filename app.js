const express = require('express');
const bodyParser = require('body-parser');
const ExcelJS = require('exceljs');
const fs = require('fs');
const path = require('path');

const app = express();
const PORT = 3000;
const BOX_FOLDER = 'C:/Users/K230086/Box/作_03_81_01_原燃利活用簡易倉庫II期/10_Ⅱ期工事/98.個人フォルダ/09.萩原/アンケート集計';
const TEMPLATE_PATH = path.join(__dirname, 'template', '熱中症対策チェックシート.xlsx');

app.use(bodyParser.urlencoded({ extended: true }));
app.use(express.static('public'));

app.post('/submit', async (req, res) => {
  const data = req.body;

  const safeValue = (v) => {
    if (Array.isArray(v)) return v.length > 0 ? v[0] : '';
    if (typeof v === 'string') return v.trim();
    return '';
  };

  const company = safeValue(data.company);
  const date = new Date().toISOString().split('T')[0];
  const time = new Date().toLocaleTimeString('ja-JP', { hour12: false }).replace(/:/g, '');
  const fileName = `${date}_${company}_${time}.xlsx`;
  const OUTPUT_PATH = path.join(BOX_FOLDER, fileName);

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(TEMPLATE_PATH);
  const worksheet = workbook.getWorksheet('熱中症対策チェックシート');
  if (!worksheet) return res.status(500).send('テンプレートシートが見つかりません');

  worksheet.getCell('BG2').value = company;
  worksheet.getCell('BG4').value = new Date().toLocaleDateString('ja-JP');

  let tools = data.tools;
  if (!Array.isArray(tools)) tools = tools ? [tools] : [];
  worksheet.getCell('W4').value = tools.length > 0 ? tools.join('、') : '';

  // 1人目 基本情報
  worksheet.getCell('S6').value = safeValue(data.name1);
  worksheet.getCell('H4').value = safeValue(data.name1);
  worksheet.getCell('S9').value = safeValue(data.sleep1);
  worksheet.getCell('S11').value = safeValue(data.fatigue1);
  worksheet.getCell('S13').value = safeValue(data.morning1);
  worksheet.getCell('S15').value = safeValue(data.Hangover1);
  worksheet.getCell('S17').value = safeValue(data.condition1);
  worksheet.getCell('S19').value = safeValue(data.clothing1);
  worksheet.getCell('S21').value = safeValue(data.doctor1);
  worksheet.getCell('S28').value = safeValue(data.heat1);
  worksheet.getCell('S35').value = safeValue(data.possible1);
  worksheet.getCell('S53').value = safeValue(data.lunch1);
  worksheet.getCell('S55').value = safeValue(data.water1);
  worksheet.getCell('S57').value = safeValue(data.Complexion1);
  worksheet.getCell('S59').value = safeValue(data.work1);

  // 1人目 摂取状況
  worksheet.getCell('S38').value = safeValue(data['intake_08301']);
  worksheet.getCell('S40').value = safeValue(data['intake_09001']);
  worksheet.getCell('S42').value = safeValue(data['intake_09301']);
  worksheet.getCell('S44').value = safeValue(data['intake_10001']);
  worksheet.getCell('S46').value = safeValue(data['intake_10301']);
  worksheet.getCell('S48').value = safeValue(data['intake_11001']);
  worksheet.getCell('S50').value = safeValue(data['intake_11301']);

  worksheet.getCell('S62').value = safeValue(data['intake_13001']);
  worksheet.getCell('S64').value = safeValue(data['intake_14001']);
  worksheet.getCell('S66').value = safeValue(data['intake_14301']);
  worksheet.getCell('S68').value = safeValue(data['intake_15001']);
  worksheet.getCell('S70').value = safeValue(data['intake_15301']);
  worksheet.getCell('S72').value = safeValue(data['intake_16001']);
  worksheet.getCell('S74').value = safeValue(data['intake_16301']);

  // 2人目 基本情報
  worksheet.getCell('W6').value = safeValue(data.name2);
  worksheet.getCell('W9').value = safeValue(data.sleep2);
  worksheet.getCell('W11').value = safeValue(data.fatigue2);
  worksheet.getCell('W13').value = safeValue(data.morning2);
  worksheet.getCell('W15').value = safeValue(data.Hangover2);
  worksheet.getCell('W17').value = safeValue(data.condition2);
  worksheet.getCell('W19').value = safeValue(data.clothing2);
  worksheet.getCell('W21').value = safeValue(data.doctor2);
  worksheet.getCell('W28').value = safeValue(data.heat2);
  worksheet.getCell('W35').value = safeValue(data.possible2);
  worksheet.getCell('W53').value = safeValue(data.lunch2);
  worksheet.getCell('W55').value = safeValue(data.water2);
  worksheet.getCell('W57').value = safeValue(data.Complexion2);
  worksheet.getCell('W59').value = safeValue(data.work2);

  // 2人目 摂取状況
  worksheet.getCell('W38').value = safeValue(data['intake_08302']);
  worksheet.getCell('W40').value = safeValue(data['intake_09002']);
  worksheet.getCell('W42').value = safeValue(data['intake_09302']);
  worksheet.getCell('W44').value = safeValue(data['intake_10002']);
  worksheet.getCell('W46').value = safeValue(data['intake_10302']);
  worksheet.getCell('W48').value = safeValue(data['intake_11002']);
  worksheet.getCell('W50').value = safeValue(data['intake_11302']);

  worksheet.getCell('W62').value = safeValue(data['intake_13002']);
  worksheet.getCell('W64').value = safeValue(data['intake_14002']);
  worksheet.getCell('W66').value = safeValue(data['intake_14302']);
  worksheet.getCell('W68').value = safeValue(data['intake_15002']);
  worksheet.getCell('W70').value = safeValue(data['intake_15302']);
  worksheet.getCell('W72').value = safeValue(data['intake_16002']);
  worksheet.getCell('W74').value = safeValue(data['intake_16302']);

  // 3人目 基本情報
  worksheet.getCell('AA6').value = safeValue(data.name3);
  worksheet.getCell('AA9').value = safeValue(data.sleep3);
  worksheet.getCell('AA11').value = safeValue(data.fatigue3);
  worksheet.getCell('AA13').value = safeValue(data.morning3);
  worksheet.getCell('AA15').value = safeValue(data.Hangover3);
  worksheet.getCell('AA17').value = safeValue(data.condition3);
  worksheet.getCell('AA19').value = safeValue(data.clothing3);
  worksheet.getCell('AA21').value = safeValue(data.doctor3);
  worksheet.getCell('AA28').value = safeValue(data.heat3);
  worksheet.getCell('AA35').value = safeValue(data.possible3);
  worksheet.getCell('AA53').value = safeValue(data.lunch3);
  worksheet.getCell('AA55').value = safeValue(data.water3);
  worksheet.getCell('AA57').value = safeValue(data.Complexion3);
  worksheet.getCell('AA59').value = safeValue(data.work3);

  // 3人目 摂取状況
  worksheet.getCell('AA38').value = safeValue(data['intake_08303']);
  worksheet.getCell('AA40').value = safeValue(data['intake_09003']);
  worksheet.getCell('AA42').value = safeValue(data['intake_09303']);
  worksheet.getCell('AA44').value = safeValue(data['intake_10003']);
  worksheet.getCell('AA46').value = safeValue(data['intake_10303']);
  worksheet.getCell('AA48').value = safeValue(data['intake_11003']);
  worksheet.getCell('AA50').value = safeValue(data['intake_11303']);

  worksheet.getCell('AA62').value = safeValue(data['intake_13003']);
  worksheet.getCell('AA64').value = safeValue(data['intake_14003']);
  worksheet.getCell('AA66').value = safeValue(data['intake_14303']);
  worksheet.getCell('AA68').value = safeValue(data['intake_15003']);
  worksheet.getCell('AA70').value = safeValue(data['intake_15303']);
  worksheet.getCell('AA72').value = safeValue(data['intake_16003']);
  worksheet.getCell('AA74').value = safeValue(data['intake_16303']);

  // 4人目 基本情報
  worksheet.getCell('AE6').value = safeValue(data.name4);
  worksheet.getCell('AE9').value = safeValue(data.sleep4);
  worksheet.getCell('AE11').value = safeValue(data.fatigue4);
  worksheet.getCell('AE13').value = safeValue(data.morning4);
  worksheet.getCell('AE15').value = safeValue(data.Hangover4);
  worksheet.getCell('AE17').value = safeValue(data.condition4);
  worksheet.getCell('AE19').value = safeValue(data.clothing4);
  worksheet.getCell('AE21').value = safeValue(data.doctor4);
  worksheet.getCell('AE28').value = safeValue(data.heat4);
  worksheet.getCell('AE35').value = safeValue(data.possible4);
  worksheet.getCell('AE53').value = safeValue(data.lunch4);
  worksheet.getCell('AE55').value = safeValue(data.water4);
  worksheet.getCell('AE57').value = safeValue(data.Complexion4);
  worksheet.getCell('AE59').value = safeValue(data.work4);

  // 4人目 摂取状況
  worksheet.getCell('AE38').value = safeValue(data['intake_08304']);
  worksheet.getCell('AE40').value = safeValue(data['intake_09004']);
  worksheet.getCell('AE42').value = safeValue(data['intake_09304']);
  worksheet.getCell('AE44').value = safeValue(data['intake_10004']);
  worksheet.getCell('AE46').value = safeValue(data['intake_10304']);
  worksheet.getCell('AE48').value = safeValue(data['intake_11004']);
  worksheet.getCell('AE50').value = safeValue(data['intake_11304']);

  worksheet.getCell('AE62').value = safeValue(data['intake_13004']);
  worksheet.getCell('AE64').value = safeValue(data['intake_14004']);
  worksheet.getCell('AE66').value = safeValue(data['intake_14304']);
  worksheet.getCell('AE68').value = safeValue(data['intake_15004']);
  worksheet.getCell('AE70').value = safeValue(data['intake_15304']);
  worksheet.getCell('AE72').value = safeValue(data['intake_16004']);
  worksheet.getCell('AE74').value = safeValue(data['intake_16304']);

  // 5人目 基本情報
  worksheet.getCell('AI6').value = safeValue(data.name5);
  worksheet.getCell('AI9').value = safeValue(data.sleep5);
  worksheet.getCell('AI11').value = safeValue(data.fatigue5);
  worksheet.getCell('AI13').value = safeValue(data.morning5);
  worksheet.getCell('AI15').value = safeValue(data.Hangover5);
  worksheet.getCell('AI17').value = safeValue(data.condition5);
  worksheet.getCell('AI19').value = safeValue(data.clothing5);
  worksheet.getCell('AI21').value = safeValue(data.doctor5);
  worksheet.getCell('AI28').value = safeValue(data.heat5);
  worksheet.getCell('AI35').value = safeValue(data.possible5);
  worksheet.getCell('AI53').value = safeValue(data.lunch5);
  worksheet.getCell('AI55').value = safeValue(data.water5);
  worksheet.getCell('AI57').value = safeValue(data.Complexion5);
  worksheet.getCell('AI59').value = safeValue(data.work5);

  // 5人目 摂取状況
  worksheet.getCell('AI38').value = safeValue(data['intake_08305']);
  worksheet.getCell('AI40').value = safeValue(data['intake_09005']);
  worksheet.getCell('AI42').value = safeValue(data['intake_09305']);
  worksheet.getCell('AI44').value = safeValue(data['intake_10005']);
  worksheet.getCell('AI46').value = safeValue(data['intake_10305']);
  worksheet.getCell('AI48').value = safeValue(data['intake_11005']);
  worksheet.getCell('AI50').value = safeValue(data['intake_11305']);

  worksheet.getCell('AI62').value = safeValue(data['intake_13005']);
  worksheet.getCell('AI64').value = safeValue(data['intake_14005']);
  worksheet.getCell('AI66').value = safeValue(data['intake_14305']);
  worksheet.getCell('AI68').value = safeValue(data['intake_15005']);
  worksheet.getCell('AI70').value = safeValue(data['intake_15305']);
  worksheet.getCell('AI72').value = safeValue(data['intake_16005']);
  worksheet.getCell('AI74').value = safeValue(data['intake_16305']);

  // 6人目 基本情報
  worksheet.getCell('AM6').value = safeValue(data.name6);
  worksheet.getCell('AM9').value = safeValue(data.sleep6);
  worksheet.getCell('AM11').value = safeValue(data.fatigue6);
  worksheet.getCell('AM13').value = safeValue(data.morning6);
  worksheet.getCell('AM15').value = safeValue(data.Hangover6);
  worksheet.getCell('AM17').value = safeValue(data.condition6);
  worksheet.getCell('AM19').value = safeValue(data.clothing6);
  worksheet.getCell('AM21').value = safeValue(data.doctor6);
  worksheet.getCell('AM28').value = safeValue(data.heat6);
  worksheet.getCell('AM35').value = safeValue(data.possible6);
  worksheet.getCell('AM53').value = safeValue(data.lunch6);
  worksheet.getCell('AM55').value = safeValue(data.water6);
  worksheet.getCell('AM57').value = safeValue(data.Complexion6);
  worksheet.getCell('AM59').value = safeValue(data.work6);

  // 6人目 摂取状況
  worksheet.getCell('AM38').value = safeValue(data['intake_08306']);
  worksheet.getCell('AM40').value = safeValue(data['intake_09006']);
  worksheet.getCell('AM42').value = safeValue(data['intake_09306']);
  worksheet.getCell('AM44').value = safeValue(data['intake_10006']);
  worksheet.getCell('AM46').value = safeValue(data['intake_10306']);
  worksheet.getCell('AM48').value = safeValue(data['intake_11006']);
  worksheet.getCell('AM50').value = safeValue(data['intake_11306']);

  worksheet.getCell('AM62').value = safeValue(data['intake_13006']);
  worksheet.getCell('AM64').value = safeValue(data['intake_14006']);
  worksheet.getCell('AM66').value = safeValue(data['intake_14306']);
  worksheet.getCell('AM68').value = safeValue(data['intake_15006']);
  worksheet.getCell('AM70').value = safeValue(data['intake_15306']);
  worksheet.getCell('AM72').value = safeValue(data['intake_16006']);
  worksheet.getCell('AM74').value = safeValue(data['intake_16306']);

  // 7人目 基本情報
  worksheet.getCell('AQ6').value = safeValue(data.name7);
  worksheet.getCell('AQ9').value = safeValue(data.sleep7);
  worksheet.getCell('AQ11').value = safeValue(data.fatigue7);
  worksheet.getCell('AQ13').value = safeValue(data.morning7);
  worksheet.getCell('AQ15').value = safeValue(data.Hangover7);
  worksheet.getCell('AQ17').value = safeValue(data.condition7);
  worksheet.getCell('AQ19').value = safeValue(data.clothing7);
  worksheet.getCell('AQ21').value = safeValue(data.doctor7);
  worksheet.getCell('AQ28').value = safeValue(data.heat7);
  worksheet.getCell('AQ35').value = safeValue(data.possible7);
  worksheet.getCell('AQ53').value = safeValue(data.lunch7);
  worksheet.getCell('AQ55').value = safeValue(data.water7);
  worksheet.getCell('AQ57').value = safeValue(data.Complexion7);
  worksheet.getCell('AQ59').value = safeValue(data.work7);

  // 7人目 摂取状況
  worksheet.getCell('AQ38').value = safeValue(data['intake_08307']);
  worksheet.getCell('AQ40').value = safeValue(data['intake_09007']);
  worksheet.getCell('AQ42').value = safeValue(data['intake_09307']);
  worksheet.getCell('AQ44').value = safeValue(data['intake_10007']);
  worksheet.getCell('AQ46').value = safeValue(data['intake_10307']);
  worksheet.getCell('AQ48').value = safeValue(data['intake_11007']);
  worksheet.getCell('AQ50').value = safeValue(data['intake_11307']);

  worksheet.getCell('AQ62').value = safeValue(data['intake_13007']);
  worksheet.getCell('AQ64').value = safeValue(data['intake_14007']);
  worksheet.getCell('AQ66').value = safeValue(data['intake_14307']);
  worksheet.getCell('AQ68').value = safeValue(data['intake_15007']);
  worksheet.getCell('AQ70').value = safeValue(data['intake_15307']);
  worksheet.getCell('AQ72').value = safeValue(data['intake_16007']);
  worksheet.getCell('AQ74').value = safeValue(data['intake_16307']);

  // 8人目 基本情報
  worksheet.getCell('AU6').value = safeValue(data.name8);
  worksheet.getCell('AU9').value = safeValue(data.sleep8);
  worksheet.getCell('AU11').value = safeValue(data.fatigue8);
  worksheet.getCell('AU13').value = safeValue(data.morning8);
  worksheet.getCell('AU15').value = safeValue(data.Hangover8);
  worksheet.getCell('AU17').value = safeValue(data.condition8);
  worksheet.getCell('AU19').value = safeValue(data.clothing8);
  worksheet.getCell('AU21').value = safeValue(data.doctor8);
  worksheet.getCell('AU28').value = safeValue(data.heat8);
  worksheet.getCell('AU35').value = safeValue(data.possible8);
  worksheet.getCell('AU53').value = safeValue(data.lunch8);
  worksheet.getCell('AU55').value = safeValue(data.water8);
  worksheet.getCell('AU57').value = safeValue(data.Complexion8);
  worksheet.getCell('AU59').value = safeValue(data.work8);

  // 8人目 摂取状況
  worksheet.getCell('AU38').value = safeValue(data['intake_08308']);
  worksheet.getCell('AU40').value = safeValue(data['intake_09008']);
  worksheet.getCell('AU42').value = safeValue(data['intake_09308']);
  worksheet.getCell('AU44').value = safeValue(data['intake_10008']);
  worksheet.getCell('AU46').value = safeValue(data['intake_10308']);
  worksheet.getCell('AU48').value = safeValue(data['intake_11008']);
  worksheet.getCell('AU50').value = safeValue(data['intake_11308']);

  worksheet.getCell('AU62').value = safeValue(data['intake_13008']);
  worksheet.getCell('AU64').value = safeValue(data['intake_14008']);
  worksheet.getCell('AU66').value = safeValue(data['intake_14308']);
  worksheet.getCell('AU68').value = safeValue(data['intake_15008']);
  worksheet.getCell('AU70').value = safeValue(data['intake_15308']);
  worksheet.getCell('AU72').value = safeValue(data['intake_16008']);
  worksheet.getCell('AU74').value = safeValue(data['intake_16308']);

  // 9人目 基本情報
  worksheet.getCell('AY6').value = safeValue(data.name9);
  worksheet.getCell('AY9').value = safeValue(data.sleep9);
  worksheet.getCell('AY11').value = safeValue(data.fatigue9);
  worksheet.getCell('AY13').value = safeValue(data.morning9);
  worksheet.getCell('AY15').value = safeValue(data.Hangover9);
  worksheet.getCell('AY17').value = safeValue(data.condition9);
  worksheet.getCell('AY19').value = safeValue(data.clothing9);
  worksheet.getCell('AY21').value = safeValue(data.doctor9);
  worksheet.getCell('AY28').value = safeValue(data.heat9);
  worksheet.getCell('AY35').value = safeValue(data.possible9);
  worksheet.getCell('AY53').value = safeValue(data.lunch9);
  worksheet.getCell('AY55').value = safeValue(data.water9);
  worksheet.getCell('AY57').value = safeValue(data.Complexion9);
  worksheet.getCell('AY59').value = safeValue(data.work9);

  // 9人目 摂取状況
  worksheet.getCell('AY38').value = safeValue(data['intake_08309']);
  worksheet.getCell('AY40').value = safeValue(data['intake_09009']);
  worksheet.getCell('AY42').value = safeValue(data['intake_09309']);
  worksheet.getCell('AY44').value = safeValue(data['intake_10009']);
  worksheet.getCell('AY46').value = safeValue(data['intake_10309']);
  worksheet.getCell('AY48').value = safeValue(data['intake_11009']);
  worksheet.getCell('AY50').value = safeValue(data['intake_11309']);

  worksheet.getCell('AY62').value = safeValue(data['intake_13009']);
  worksheet.getCell('AY64').value = safeValue(data['intake_14009']);
  worksheet.getCell('AY66').value = safeValue(data['intake_14309']);
  worksheet.getCell('AY68').value = safeValue(data['intake_15009']);
  worksheet.getCell('AY70').value = safeValue(data['intake_15309']);
  worksheet.getCell('AY72').value = safeValue(data['intake_16009']);
  worksheet.getCell('AY74').value = safeValue(data['intake_16309']);

  // 10人目 基本情報
  worksheet.getCell('BC6').value = safeValue(data.name10);
  worksheet.getCell('BC9').value = safeValue(data.sleep10);
  worksheet.getCell('BC11').value = safeValue(data.fatigue10);
  worksheet.getCell('BC13').value = safeValue(data.morning10);
  worksheet.getCell('BC15').value = safeValue(data.Hangover10);
  worksheet.getCell('BC17').value = safeValue(data.condition10);
  worksheet.getCell('BC19').value = safeValue(data.clothing10);
  worksheet.getCell('BC21').value = safeValue(data.doctor10);
  worksheet.getCell('BC28').value = safeValue(data.heat10);
  worksheet.getCell('BC35').value = safeValue(data.possible10);
  worksheet.getCell('BC53').value = safeValue(data.lunch10);
  worksheet.getCell('BC55').value = safeValue(data.water10);
  worksheet.getCell('BC57').value = safeValue(data.Complexion10);
  worksheet.getCell('BC59').value = safeValue(data.work10);

  // 10人目 摂取状況
  worksheet.getCell('BC38').value = safeValue(data['intake_083010']);
  worksheet.getCell('BC40').value = safeValue(data['intake_090010']);
  worksheet.getCell('BC42').value = safeValue(data['intake_093010']);
  worksheet.getCell('BC44').value = safeValue(data['intake_100010']);
  worksheet.getCell('BC46').value = safeValue(data['intake_103010']);
  worksheet.getCell('BC48').value = safeValue(data['intake_110010']);
  worksheet.getCell('BC50').value = safeValue(data['intake_113010']);

  worksheet.getCell('BC62').value = safeValue(data['intake_130010']);
  worksheet.getCell('BC64').value = safeValue(data['intake_140010']);
  worksheet.getCell('BC66').value = safeValue(data['intake_143010']);
  worksheet.getCell('BC68').value = safeValue(data['intake_150010']);
  worksheet.getCell('BC70').value = safeValue(data['intake_153010']);
  worksheet.getCell('BC72').value = safeValue(data['intake_160010']);
  worksheet.getCell('BC74').value = safeValue(data['intake_163010']);

  await workbook.xlsx.writeFile(OUTPUT_PATH);
  res.send('保存が完了しました（全項目空欄対応）');
});

app.listen(PORT, () => {
  console.log(`Server running at http://localhost:${PORT}`);
});
