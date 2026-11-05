// Рендер ролика: кадр за кадром через seek(t), JPEG-поток в ffmpeg
const { chromium } = require('playwright');
const { spawn } = require('child_process');
(async () => {
  const FPS = 30, from = +(process.argv[2] || 0), to = +(process.argv[3] || 54), out = process.argv[4] || 'video.mp4';
  const b = await chromium.launch(); const p = await b.newPage({ viewport: { width: 1920, height: 1080 } });
  await p.goto('file://' + __dirname + '/index.html'); await p.waitForFunction(() => window.__ready);
  const ff = spawn('ffmpeg', ['-y', '-loglevel', 'error', '-f', 'image2pipe', '-framerate', String(FPS), '-c:v', 'mjpeg', '-i', '-', '-c:v', 'libx264', '-preset', 'veryfast', '-crf', '17', '-pix_fmt', 'yuv420p', out]);
  const n0 = Math.round(from * FPS), n1 = Math.round(to * FPS);
  for (let n = n0; n < n1; n++) {
    await p.evaluate(t => seek(t), n / FPS);
    const buf = await p.screenshot({ type: 'jpeg', quality: 92 });
    if (!ff.stdin.write(buf)) await new Promise(r => ff.stdin.once('drain', r));
    if (n % 150 === 0) console.log('frame', n, '/', n1);
  }
  ff.stdin.end(); await new Promise(r => ff.on('close', r)); await b.close(); console.log('done', out);
})();
