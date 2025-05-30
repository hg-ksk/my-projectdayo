const QRCode = require('qrcode');

const URL = 'http://localhost:3000'; // 公開時はURLを差し替える

QRCode.toFile('qrcode.png', URL, { width: 300 }, (err) => {
  if (err) throw err;
  console.log('QRコードを qrcode.png に保存しました。');
});
