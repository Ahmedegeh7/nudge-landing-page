const { chromium } = require('playwright');
const fs = require('fs');

const FONT_LINK = `<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Anton&display=swap">`;

// wordmark lockups
function wordmarkHTML(color, dotColor) {
  return `<!doctype html><html><head><meta charset="utf-8">${FONT_LINK}
  <style>
    html,body{margin:0;padding:0;background:transparent;}
    .lockup{
      display:inline-flex; align-items:baseline;
      font-family:'Anton', sans-serif;
      font-size:160px;
      color:${color};
      letter-spacing:1px;
      padding:40px 50px;
    }
    .dot{ color:${dotColor}; display:inline-block; transform:translateY(-14px) rotate(-18deg); margin-left:2px; }
  </style></head>
  <body><div class="lockup">Nudge<span class="dot">.</span></div></body></html>`;
}

// icon-only mark ("N.") for favicon / app icon, on a rounded square tile
function iconHTML(bg, fg, dot) {
  return `<!doctype html><html><head><meta charset="utf-8">${FONT_LINK}
  <style>
    html,body{margin:0;padding:0;}
    .tile{
      width:512px; height:512px;
      background:${bg};
      border-radius:${bg === 'transparent' ? '0' : '112px'};
      display:flex; align-items:center; justify-content:center;
      font-family:'Anton', sans-serif;
    }
    .n{ font-size:300px; color:${fg}; line-height:1; }
    .dot{ color:${dot}; display:inline-block; transform:translateY(-30px) rotate(-18deg); }
  </style></head>
  <body><div class="tile"><span class="n">N<span class="dot">.</span></span></div></body></html>`;
}

const INK = '#1E2A2F';
const PAPER = '#EDE3CF';
const ACCENT = '#E2531D';

async function shot(html, path, width, height) {
  const browser = await chromium.launch({ executablePath: '/opt/pw-browsers/chromium' });
  const page = await browser.newPage({ viewport: { width, height }, deviceScaleFactor: 2 });
  await page.setContent(html, { waitUntil: 'networkidle' });
  await page.waitForTimeout(300);
  const el = await page.$('body > *');
  await el.screenshot({ path, omitBackground: true });
  await browser.close();
}

(async () => {
  // Wordmark, ink-on-transparent (for light backgrounds)
  await shot(wordmarkHTML(INK, ACCENT), 'nudge-wordmark-ink.png', 620, 260);
  // Wordmark, paper-on-transparent (for dark backgrounds)
  await shot(wordmarkHTML(PAPER, ACCENT), 'nudge-wordmark-reversed.png', 620, 260);
  // Icon tile, ink background, paper "n", accent dot (app icon style, filled square)
  await shot(iconHTML(INK, PAPER, ACCENT), 'nudge-icon-tile.png', 512, 512);
  // Icon, transparent background, ink "n", accent dot (for compositing anywhere)
  await shot(iconHTML('transparent', INK, ACCENT), 'nudge-icon-transparent.png', 512, 512);
  console.log('done');
})();
