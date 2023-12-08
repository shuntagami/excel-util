const fs = require('fs');
const path = require('path');

const sourcePath = '100x100.jpeg'; // 元の画像ファイルのパス
const targetDir = './images'; // 生成するファイルの保存先ディレクトリ

for (let i = 0; i <= 1000; i++) {
  const targetPath = path.join(targetDir, `image-${i}.jpeg`);
  fs.copyFileSync(sourcePath, targetPath);
}
