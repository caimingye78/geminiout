import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';
import fs from 'node:fs';

const chromePath = [
  '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome',
  '/Applications/Chromium.app/Contents/MacOS/Chromium',
  '/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge',
].find((candidate) => fs.existsSync(candidate));

const localReadCache = new Map<string, { title: string; body: string; url: string }>();

function cleanGeminiText(text: string) {
  const drop = new Set([
    '关于 Gemini',
    '获取 Gemini 应用',
    '订阅',
    '企业应用场景',
    '登录',
    '在新窗口中打开',
    '《Google 隐私权政策》',
    'Google 服务条款',
    '你的隐私权与 Gemini 应用',
  ]);
  const lines = text
    .split('\n')
    .map((line) => line.trim())
    .filter(Boolean)
    .filter((line) => !drop.has(line));
  const footerIndex = lines.findIndex((line) => line.includes('Gemini 显示的信息'));
  return (footerIndex >= 0 ? lines.slice(0, footerIndex) : lines).join('\n').trim();
}

async function readGeminiShare(url: string) {
  if (localReadCache.has(url)) return localReadCache.get(url);
  if (!chromePath) throw new Error('没有找到 Chrome / Chromium / Edge，无法本地读取 Gemini 分享页');
  if (!/^https:\/\/gemini\.google\.com\/share\/[A-Za-z0-9_-]+$/.test(url)) {
    throw new Error('只支持 gemini.google.com/share 公开链接');
  }

  const { chromium } = await import('playwright-core');
  const browser = await chromium.launch({ executablePath: chromePath, headless: true });
  try {
    const page = await browser.newPage({ viewport: { width: 1440, height: 1200 } });
    await page.goto(url, { waitUntil: 'domcontentloaded', timeout: 45000 });
    await page.waitForTimeout(7000);
    const title = (await page.locator('h1').first().innerText({ timeout: 5000 }).catch(() => '')).trim();
    const rawText = await page.locator('body').innerText({ timeout: 10000 }).catch(() => '');
    const body = cleanGeminiText(rawText)
      .split('\n')
      .filter((line) => line !== url)
      .join('\n')
      .trim();
    if (body.length < 20) throw new Error('没有读到正文，可能需要登录、链接失效或页面尚未公开');
    const result = { title, body, url };
    localReadCache.set(url, result);
    return result;
  } finally {
    await browser.close();
  }
}

// https://vite.dev/config/
export default defineConfig({
  base: '/geminiout/',
  plugins: [
    react(),
    {
      name: 'local-gemini-reader',
      configureServer(server) {
        server.middlewares.use('/geminiout/api/read-gemini', async (req, res) => {
          try {
            const requestUrl = new URL(req.url ?? '', 'http://localhost');
            const target = requestUrl.searchParams.get('url') ?? '';
            const result = await readGeminiShare(target);
            res.statusCode = 200;
            res.setHeader('Content-Type', 'application/json; charset=utf-8');
            res.end(JSON.stringify(result));
          } catch (error) {
            res.statusCode = 500;
            res.setHeader('Content-Type', 'application/json; charset=utf-8');
            res.end(
              JSON.stringify({
                error: error instanceof Error ? error.message : '未知错误',
              }),
            );
          }
        });
      },
    },
  ],
});
