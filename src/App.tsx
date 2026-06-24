import { useMemo, useRef, useState } from 'react';
import MarkdownIt from 'markdown-it';
import katex from 'katex';
import { saveAs } from 'file-saver';
import {
  AlignmentType,
  BorderStyle,
  Document,
  HeadingLevel,
  Packer,
  Paragraph,
  ShadingType,
  Table,
  TableCell,
  TableRow,
  TextRun,
  WidthType,
} from 'docx';
import {
  Archive,
  CheckCircle2,
  CloudDownload,
  Download,
  FileDown,
  FileText,
  FolderOpen,
  ListChecks,
  Link2,
  MessageSquareText,
  ScanText,
  Search,
  Settings2,
  Sparkles,
  UploadCloud,
} from 'lucide-react';
import 'katex/dist/katex.min.css';
import './App.css';

type Role = 'user' | 'assistant' | 'system' | 'unknown';

type ChatMessage = {
  id: string;
  role: Role;
  text: string;
  createdAt?: string;
};

type Conversation = {
  id: string;
  title: string;
  messages: ChatMessage[];
  sourceName: string;
  createdAt?: string;
};

type ExportStyle = 'editorial' | 'compact' | 'academic';

type LinkFetchState = {
  status: 'loading' | 'success' | 'failed';
  title?: string;
  chars?: number;
  error?: string;
};

const roleLabels: Record<Role, string> = {
  user: '你',
  assistant: 'Gemini',
  system: '系统',
  unknown: '消息',
};

const md = new MarkdownIt({
  html: false,
  linkify: true,
  breaks: true,
  typographer: true,
});

const READER_PROXY_PREFIX = 'https://r.jina.ai/http://r.jina.ai/http://';

const sampleText = `# 一个含公式的 Gemini 对话示例

**你：** 请解释二次公式。

**Gemini：** 对于方程 $ax^2 + bx + c = 0$，解为：

$$x = \\frac{-b \\pm \\sqrt{b^2 - 4ac}}{2a}$$

当判别式 $\\Delta=b^2-4ac$ 大于 0 时，有两个不同实根。

| 情况 | 根 |
| --- | --- |
| $\\Delta > 0$ | 两个实根 |
| $\\Delta = 0$ | 一个重根 |
| $\\Delta < 0$ | 共轭复根 |`;

function App() {
  const inputRef = useRef<HTMLInputElement | null>(null);
  const screenshotRef = useRef<HTMLInputElement | null>(null);
  const [conversations, setConversations] = useState<Conversation[]>([]);
  const [selectedIds, setSelectedIds] = useState<Set<string>>(new Set());
  const [detectedLinks, setDetectedLinks] = useState<string[]>([]);
  const [linkFetches, setLinkFetches] = useState<Record<string, LinkFetchState>>({});
  const [isFetchingLinks, setIsFetchingLinks] = useState(false);
  const [pastedLinkText, setPastedLinkText] = useState('');
  const [activeId, setActiveId] = useState<string | null>(null);
  const [query, setQuery] = useState('');
  const [style, setStyle] = useState<ExportStyle>('editorial');
  const [includeMeta, setIncludeMeta] = useState(true);
  const [status, setStatus] = useState('等待导入 Gemini 聊天记录');

  const activeConversation = conversations.find((item) => item.id === activeId) ?? conversations[0];
  const selectedConversations = conversations.filter((item) => selectedIds.has(item.id));
  const visibleConversations = conversations.filter((item) => {
    const haystack = `${item.title} ${item.messages.map((message) => message.text).join(' ')}`.toLowerCase();
    return haystack.includes(query.trim().toLowerCase());
  });

  const stats = useMemo(() => {
    const messageCount = conversations.reduce((sum, item) => sum + item.messages.length, 0);
    const wordCount = conversations.reduce(
      (sum, item) => sum + item.messages.reduce((inner, message) => inner + message.text.length, 0),
      0,
    );
    return { messageCount, wordCount };
  }, [conversations]);

  async function handleFiles(files: FileList | null) {
    if (!files?.length) return;
    setStatus('正在解析文件...');
    const parsed: Conversation[] = [];
    for (const file of Array.from(files)) {
      const text = await file.text();
      parsed.push(...parseFile(file.name, text));
    }
    const deduped = dedupeConversations(parsed);
    setConversations(deduped);
    setSelectedIds(new Set(deduped.map((item) => item.id)));
    setActiveId(deduped[0]?.id ?? null);
    setStatus(`已解析 ${deduped.length} 个会话，可以导出 Word`);
  }

  async function handleScreenshot(files: FileList | null) {
    const file = files?.[0];
    if (!file) return;
    setStatus('正在 OCR 识别截图里的链接...');
    try {
      const { recognize } = await import('tesseract.js');
      const result = await recognize(file, 'eng', {
        logger: (message) => {
          if (message.status === 'recognizing text') {
            setStatus(`正在 OCR 识别截图... ${Math.round(message.progress * 100)}%`);
          }
        },
      });
      const links = extractGeminiLinks(result.data.text);
      addDetectedLinks(links);
      setStatus(links.length ? `从截图识别出 ${links.length} 个 Gemini 链接` : '没有识别到完整 Gemini 分享链接');
    } catch (error) {
      setStatus(`截图识别失败：${error instanceof Error ? error.message : '未知错误'}`);
    } finally {
      if (screenshotRef.current) screenshotRef.current.value = '';
    }
  }

  function addDetectedLinks(links: string[]) {
    if (!links.length) return;
    setDetectedLinks((current) => Array.from(new Set([...current, ...links])));
  }

  function extractLinksFromPaste() {
    const links = extractGeminiLinks(pastedLinkText);
    addDetectedLinks(links);
    setStatus(links.length ? `从粘贴文本提取出 ${links.length} 个 Gemini 链接` : '粘贴内容里没有找到 Gemini 分享链接');
  }

  async function fetchDetectedLinkContent() {
    if (!detectedLinks.length || isFetchingLinks) return;
    setIsFetchingLinks(true);
    const fetched: Conversation[] = [];
    let failed = 0;

    for (const [index, link] of detectedLinks.entries()) {
      setStatus(`正在拉取公开链接 ${index + 1}/${detectedLinks.length}...`);
      setLinkFetches((current) => ({ ...current, [link]: { status: 'loading' } }));
      try {
        const response = await fetch(`${READER_PROXY_PREFIX}${link}`);
        if (!response.ok) throw new Error(`HTTP ${response.status}`);
        const text = await response.text();
        const parsed = parseReaderMarkdown(link, text);
        if (parsed.body.length < 20) throw new Error('没有读到正文，可能需要登录或分享已失效');
        const conversation = conversationFromReader(link, parsed, index);
        fetched.push(conversation);
        setLinkFetches((current) => ({
          ...current,
          [link]: { status: 'success', title: conversation.title, chars: parsed.body.length },
        }));
      } catch (error) {
        failed += 1;
        setLinkFetches((current) => ({
          ...current,
          [link]: {
            status: 'failed',
            error: error instanceof Error ? error.message : '未知错误',
          },
        }));
      }
    }

    if (fetched.length) {
      const merged = dedupeConversations([...conversations, ...fetched]);
      setConversations(merged);
      setSelectedIds(new Set(merged.map((item) => item.id)));
      setActiveId(fetched[0].id);
    }
    setStatus(`公开链接拉取完成：成功 ${fetched.length} 个，失败 ${failed} 个`);
    setIsFetchingLinks(false);
  }

  function loadSample() {
    const demo: Conversation = {
      id: 'sample',
      title: '公式排版示例',
      sourceName: '内置示例',
      messages: [
        { id: 'sample-1', role: 'user', text: '请解释二次公式。' },
        {
          id: 'sample-2',
          role: 'assistant',
          text: sampleText.replace(/^# .+\n+/, '').replace(/\*\*你：\*\* 请解释二次公式。\n+/, ''),
        },
      ],
    };
    setConversations([demo]);
    setSelectedIds(new Set([demo.id]));
    setActiveId(demo.id);
    setStatus('已载入示例，可直接预览和导出');
  }

  async function exportDocx() {
    const targets = selectedConversations.length ? selectedConversations : conversations;
    if (!targets.length) {
      setStatus('请先导入聊天记录');
      return;
    }
    setStatus('正在生成 Word 文档...');
    const doc = buildDocx(targets, { style, includeMeta });
    const blob = await Packer.toBlob(doc);
    const date = new Date().toISOString().slice(0, 10);
    saveAs(blob, `gemini-chat-export-${date}.docx`);
    setStatus(`已导出 ${targets.length} 个会话`);
  }

  function toggleConversation(id: string) {
    setSelectedIds((current) => {
      const next = new Set(current);
      if (next.has(id)) next.delete(id);
      else next.add(id);
      return next;
    });
  }

  function selectAll() {
    setSelectedIds(new Set(visibleConversations.map((item) => item.id)));
  }

  function clearSelection() {
    setSelectedIds(new Set());
  }

  return (
    <main className="app-shell">
      <section className="sidebar import-panel">
        <div className="brand-row">
          <div className="brand-mark">
            <Archive size={22} />
          </div>
          <div>
            <p className="eyebrow">Gemini Export Studio</p>
            <h1>聊天记录转 Word</h1>
          </div>
        </div>

        <button className="drop-zone" onClick={() => inputRef.current?.click()}>
          <UploadCloud size={30} />
          <span>导入 Gemini JSON / Markdown / TXT</span>
          <small>支持多文件，内容只在本机浏览器处理</small>
        </button>
        <input
          ref={inputRef}
          className="sr-only"
          type="file"
          multiple
          accept=".json,.md,.markdown,.txt,.html"
          onChange={(event) => void handleFiles(event.target.files)}
        />

        <div className="quick-actions">
          <button type="button" onClick={loadSample}>
            <Sparkles size={16} />
            载入示例
          </button>
          <button type="button" onClick={() => screenshotRef.current?.click()}>
            <ScanText size={16} />
            识别截图
          </button>
          <button type="button" onClick={selectAll}>
            <ListChecks size={16} />
            全选结果
          </button>
        </div>
        <input
          ref={screenshotRef}
          className="sr-only"
          type="file"
          accept="image/png,image/jpeg,image/webp,image/heic,image/heif"
          onChange={(event) => void handleScreenshot(event.target.files)}
        />

        <div className="paste-link-box">
          <textarea
            value={pastedLinkText}
            onChange={(event) => setPastedLinkText(event.target.value)}
            placeholder="粘贴任意文本、书签导出片段或分享弹窗内容..."
            rows={4}
          />
          <button type="button" onClick={extractLinksFromPaste}>
            <Link2 size={16} />
            提取链接
          </button>
        </div>

        {detectedLinks.length > 0 && (
          <div className="link-results">
            <div className="section-title">
              <Link2 size={17} />
              <span>截图链接</span>
            </div>
            <button
              className="fetch-links-button"
              type="button"
              disabled={isFetchingLinks}
              onClick={() => void fetchDetectedLinkContent()}
            >
              <CloudDownload size={16} />
              {isFetchingLinks ? '正在拉取公开内容...' : '拉取公开内容'}
            </button>
            <p className="link-hint">会自动把可访问的 Gemini 分享页转成会话，失败项会保留状态。</p>
            {detectedLinks.map((link) => {
              const fetchState = linkFetches[link];
              return (
                <div key={link} className={`link-result-row ${fetchState?.status ?? ''}`}>
                  <a href={link} target="_blank" rel="noreferrer">
                    {link}
                  </a>
                  {fetchState && (
                    <small>
                      {fetchState.status === 'loading' && '拉取中'}
                      {fetchState.status === 'success' && `已加入：${fetchState.title} · ${fetchState.chars ?? 0} 字符`}
                      {fetchState.status === 'failed' && `失败：${fetchState.error}`}
                    </small>
                  )}
                </div>
              );
            })}
          </div>
        )}

        <div className="settings-block">
          <div className="section-title">
            <Settings2 size={17} />
            <span>导出样式</span>
          </div>
          <div className="segmented">
            <button className={style === 'editorial' ? 'active' : ''} onClick={() => setStyle('editorial')}>
              舒展
            </button>
            <button className={style === 'compact' ? 'active' : ''} onClick={() => setStyle('compact')}>
              紧凑
            </button>
            <button className={style === 'academic' ? 'active' : ''} onClick={() => setStyle('academic')}>
              学术
            </button>
          </div>
          <label className="toggle-row">
            <input type="checkbox" checked={includeMeta} onChange={(event) => setIncludeMeta(event.target.checked)} />
            <span>包含来源、时间和消息数</span>
          </label>
        </div>

        <div className="status-box">
          <CheckCircle2 size={18} />
          <span>{status}</span>
        </div>
      </section>

      <section className="conversation-list">
        <div className="list-header">
          <div>
            <p className="eyebrow">会话</p>
            <h2>{conversations.length || 0} 个聊天</h2>
          </div>
          <button className="icon-button" onClick={clearSelection} title="清空选择">
            <FolderOpen size={18} />
          </button>
        </div>

        <label className="search-box">
          <Search size={17} />
          <input value={query} onChange={(event) => setQuery(event.target.value)} placeholder="搜索标题或正文" />
        </label>

        <div className="metric-grid">
          <div>
            <strong>{stats.messageCount}</strong>
            <span>消息</span>
          </div>
          <div>
            <strong>{Math.round(stats.wordCount / 1000)}k</strong>
            <span>字符</span>
          </div>
        </div>

        <div className="chat-list">
          {visibleConversations.map((conversation) => (
            <button
              key={conversation.id}
              className={`chat-row ${activeConversation?.id === conversation.id ? 'active' : ''}`}
              onClick={() => setActiveId(conversation.id)}
            >
              <input
                type="checkbox"
                checked={selectedIds.has(conversation.id)}
                onChange={() => toggleConversation(conversation.id)}
                onClick={(event) => event.stopPropagation()}
              />
              <span>
                <strong>{conversation.title}</strong>
                <small>
                  {conversation.messages.length} 条消息 · {conversation.sourceName}
                </small>
              </span>
            </button>
          ))}
          {!visibleConversations.length && <p className="empty-state">还没有可显示的会话。</p>}
        </div>
      </section>

      <section className="preview-panel">
        {activeConversation ? (
          <>
            <div className="preview-header">
              <div>
                <p className="eyebrow">预览</p>
                <h2>{activeConversation.title}</h2>
              </div>
              <button className="export-button" onClick={() => void exportDocx()}>
                <FileDown size={18} />
                导出 Word
              </button>
            </div>
            <div className="document-preview">
              {activeConversation.messages.map((message) => (
                <article key={message.id} className={`message-card role-${message.role}`}>
                  <div className="message-meta">
                    <span>
                      <MessageSquareText size={15} />
                      {roleLabels[message.role]}
                    </span>
                    {message.createdAt && <time>{message.createdAt}</time>}
                  </div>
                  <RenderedMarkdown text={message.text} />
                </article>
              ))}
            </div>
          </>
        ) : (
          <div className="empty-canvas">
            <FileText size={44} />
            <h2>导入文件后在这里预览</h2>
            <p>Markdown 标记会被转成排版结构，LaTeX 公式会用 KaTeX 渲染。</p>
          </div>
        )}
      </section>

      <aside className="export-panel">
        <div className="section-title">
          <Download size={18} />
          <span>导出队列</span>
        </div>
        <strong className="queue-number">{selectedConversations.length || conversations.length}</strong>
        <p>未手动选择时默认导出全部会话。生成的 Word 会包含目录式会话分隔、角色标识、代码块、表格、引用和公式块。</p>
        <button className="wide-export" onClick={() => void exportDocx()}>
          <FileDown size={18} />
          生成 .docx
        </button>
      </aside>
    </main>
  );
}

function RenderedMarkdown({ text }: { text: string }) {
  const html = useMemo(() => renderMarkdownWithMath(text), [text]);
  return <div className="rendered-message" dangerouslySetInnerHTML={{ __html: html }} />;
}

function renderMarkdownWithMath(text: string) {
  const protectedMath: string[] = [];
  const safe = text.replace(/\$\$([\s\S]+?)\$\$|\$([^$\n]+?)\$/g, (match, block, inline) => {
    const latex = String(block ?? inline ?? '').trim();
    try {
      const html = katex.renderToString(latex, {
        displayMode: Boolean(block),
        throwOnError: false,
        output: 'html',
      });
      const token = `@@MATH_${protectedMath.length}@@`;
      protectedMath.push(`<span class="${block ? 'math-block' : 'math-inline'}">${html}</span>`);
      return token;
    } catch {
      return match;
    }
  });
  let html = md.render(safe);
  protectedMath.forEach((mathHtml, index) => {
    html = html.replace(`@@MATH_${index}@@`, mathHtml);
  });
  return html;
}

function parseFile(fileName: string, text: string): Conversation[] {
  if (fileName.toLowerCase().endsWith('.json')) {
    try {
      return parseJsonConversations(JSON.parse(text), fileName);
    } catch {
      return [plainTextConversation(fileName, text)];
    }
  }
  if (fileName.toLowerCase().endsWith('.html')) {
    const body = new DOMParser().parseFromString(text, 'text/html').body.textContent ?? text;
    return [plainTextConversation(fileName, body)];
  }
  return [plainTextConversation(fileName, text)];
}

function parseReaderMarkdown(url: string, text: string) {
  const titleMatch = text.match(/^Title:\s*(.+)$/m);
  const marker = 'Markdown Content:';
  const bodyStart = text.indexOf(marker);
  const rawBody = bodyStart >= 0 ? text.slice(bodyStart + marker.length) : text;
  const body = rawBody
    .replace(/^Title:\s*.+$/gm, '')
    .replace(/^URL Source:\s*.+$/gm, '')
    .replace(/\n{3,}/g, '\n\n')
    .trim();
  return {
    title: cleanReaderTitle(titleMatch?.[1] ?? inferTitle(body || url)),
    body,
  };
}

function cleanReaderTitle(title: string) {
  return title
    .replace(/^‎?Gemini\s*-\s*/i, '')
    .replace(/^direct access to Google AI$/i, '')
    .trim();
}

function conversationFromReader(url: string, parsed: { title: string; body: string }, index: number): Conversation {
  const title = parsed.title || `Gemini 分享 ${index + 1}`;
  return {
    id: makeId(url, parsed.body.slice(0, 120)),
    title,
    sourceName: url,
    messages: [
      {
        id: makeId(url, 'reader-message'),
        role: 'assistant',
        text: parsed.body,
      },
    ],
  };
}

function extractGeminiLinks(text: string) {
  const normalizedLines = text
    .replace(/[“”]/g, '"')
    .replace(/[，。]/g, '.')
    .replace(/[：]/g, ':')
    .split(/\n+/)
    .map((line) => line.trim())
    .filter(Boolean);
  const lineMatches = normalizedLines.flatMap((line) => {
    const normalized = line
      .replace(/gemini\.google\.com\/share\//gi, 'https://gemini.google.com/share/')
      .replace(/g\.co\/gemini\/share\//gi, 'https://g.co/gemini/share/');
    return normalized.match(/https?:\/\/(?:gemini\.google\.com\/share|g\.co\/gemini\/share)\/[A-Za-z0-9_-]+/g) ?? [];
  });
  const compact = normalizedLines
    .join('')
    .replace(/gemini\.google\.com\/share\//gi, 'https://gemini.google.com/share/')
    .replace(/g\.co\/gemini\/share\//gi, 'https://g.co/gemini/share/');
  const compactMatches = compact.match(/https?:\/\/(?:gemini\.google\.com\/share|g\.co\/gemini\/share)\/[A-Za-z0-9_-]{12}/g) ?? [];
  return Array.from(new Set([...lineMatches, ...compactMatches].map((link) => link.replace(/[.,;:，。；：]+$/, ''))));
}

function parseJsonConversations(data: unknown, sourceName: string): Conversation[] {
  const candidates = collectConversationCandidates(data);
  if (candidates.length) return candidates.map((item, index) => normalizeConversation(item, sourceName, index));
  const messages = collectMessages(data);
  if (messages.length) {
    return [
      {
        id: makeId(sourceName, 'single'),
        title: inferTitle(messages[0]?.text ?? sourceName),
        sourceName,
        messages,
      },
    ];
  }
  return [plainTextConversation(sourceName, JSON.stringify(data, null, 2))];
}

function collectConversationCandidates(value: unknown): unknown[] {
  const found: unknown[] = [];
  const visit = (node: unknown) => {
    if (!node || typeof node !== 'object') return;
    if (Array.isArray(node)) {
      node.forEach(visit);
      return;
    }
    const record = node as Record<string, unknown>;
    const possibleMessages = record.messages ?? record.turns ?? record.entries ?? record.conversation ?? record.chunks;
    if (Array.isArray(possibleMessages) && collectMessages(possibleMessages).length) {
      found.push(record);
      return;
    }
    Object.values(record).forEach(visit);
  };
  visit(value);
  return found;
}

function normalizeConversation(value: unknown, sourceName: string, index: number): Conversation {
  const record = (value && typeof value === 'object' ? value : {}) as Record<string, unknown>;
  const rawMessages = record.messages ?? record.turns ?? record.entries ?? record.conversation ?? record.chunks ?? [];
  const messages = collectMessages(rawMessages);
  const title =
    asString(record.title) ??
    asString(record.name) ??
    asString(record.conversation_title) ??
    inferTitle(messages[0]?.text ?? sourceName);
  const createdAt = asString(record.create_time) ?? asString(record.created_at) ?? asString(record.createdAt);
  return {
    id: makeId(sourceName, `${index}-${title}`),
    title,
    sourceName,
    ...(createdAt ? { createdAt } : {}),
    messages,
  };
}

function collectMessages(value: unknown): ChatMessage[] {
  const messages: ChatMessage[] = [];
  const visit = (node: unknown, indexPath: string) => {
    if (!node || typeof node !== 'object') return;
    if (Array.isArray(node)) {
      node.forEach((item, index) => visit(item, `${indexPath}-${index}`));
      return;
    }
    const record = node as Record<string, unknown>;
    const text = extractText(record);
    if (text) {
      const createdAt = asString(record.created_at) ?? asString(record.createdAt) ?? asString(record.timestamp);
      messages.push({
        id: asString(record.id) ?? `message-${indexPath}-${messages.length}`,
        role: normalizeRole(asString(record.role) ?? asString(record.author) ?? asString(record.sender)),
        text,
        ...(createdAt ? { createdAt } : {}),
      });
      return;
    }
    Object.entries(record).forEach(([key, child]) => {
      if (['metadata', 'settings', 'config'].includes(key)) return;
      visit(child, `${indexPath}-${key}`);
    });
  };
  visit(value, 'root');
  return messages.filter((message, index, list) => {
    const previous = list[index - 1];
    return !previous || previous.text !== message.text || previous.role !== message.role;
  });
}

function extractText(record: Record<string, unknown>): string | null {
  const direct =
    asString(record.text) ??
    asString(record.content) ??
    asString(record.markdown) ??
    asString(record.answer) ??
    asString(record.prompt);
  if (direct) return direct.trim();
  const parts = record.parts;
  if (Array.isArray(parts)) {
    const joined = parts.map((part) => (typeof part === 'string' ? part : extractText((part ?? {}) as Record<string, unknown>))).filter(Boolean).join('\n\n');
    return joined.trim() || null;
  }
  const content = record.content;
  if (content && typeof content === 'object') {
    const nested = content as Record<string, unknown>;
    if (Array.isArray(nested.parts)) {
      const joined = nested.parts.map((part) => (typeof part === 'string' ? part : asString((part as Record<string, unknown>)?.text))).filter(Boolean).join('\n\n');
      return joined.trim() || null;
    }
    return extractText(nested);
  }
  return null;
}

function plainTextConversation(fileName: string, text: string): Conversation {
  return {
    id: makeId(fileName, text.slice(0, 40)),
    title: inferTitle(text) || fileName,
    sourceName: fileName,
    messages: splitPlainTextMessages(text).map((message, index) => ({ ...message, id: `${fileName}-${index}` })),
  };
}

function splitPlainTextMessages(text: string): Omit<ChatMessage, 'id'>[] {
  const chunks = text
    .split(/\n(?=(?:你|用户|User|Gemini|Assistant|Model|模型|系统|System)\s*[：:])/g)
    .map((chunk) => chunk.trim())
    .filter(Boolean);
  if (chunks.length <= 1) return [{ role: 'unknown', text }];
  return chunks.map((chunk) => {
    const match = chunk.match(/^([^：:\n]{1,24})[：:]\s*([\s\S]*)$/);
    return {
      role: normalizeRole(match?.[1]),
      text: (match?.[2] ?? chunk).trim(),
    };
  });
}

function buildDocx(conversations: Conversation[], options: { style: ExportStyle; includeMeta: boolean }) {
  const spacing = options.style === 'compact' ? 150 : 220;
  const accent = options.style === 'academic' ? '334155' : '245B5B';
  const children: (Paragraph | Table)[] = [
    new Paragraph({
      text: 'Gemini 聊天记录导出',
      heading: HeadingLevel.TITLE,
      spacing: { after: 180 },
    }),
    new Paragraph({
      children: [
        new TextRun({ text: `${conversations.length} 个会话 · ${new Date().toLocaleString()}`, color: '52616B' }),
      ],
      spacing: { after: 360 },
    }),
  ];

  conversations.forEach((conversation, conversationIndex) => {
    children.push(
      new Paragraph({
        text: conversation.title,
        heading: HeadingLevel.HEADING_1,
        thematicBreak: conversationIndex > 0,
        spacing: { before: conversationIndex > 0 ? 360 : 120, after: 120 },
      }),
    );
    if (options.includeMeta) {
      children.push(
        new Paragraph({
          children: [
            new TextRun({
              text: `${conversation.sourceName} · ${conversation.messages.length} 条消息${conversation.createdAt ? ` · ${conversation.createdAt}` : ''}`,
              color: '64748B',
              size: 20,
            }),
          ],
          spacing: { after: 180 },
        }),
      );
    }
    conversation.messages.forEach((message) => {
      children.push(
        new Paragraph({
          children: [
            new TextRun({
              text: roleLabels[message.role],
              bold: true,
              color: message.role === 'assistant' ? accent : '5B341F',
            }),
            ...(message.createdAt ? [new TextRun({ text: `  ${message.createdAt}`, color: '94A3B8', size: 18 })] : []),
          ],
          spacing: { before: 160, after: 80 },
          border: {
            bottom: { style: BorderStyle.SINGLE, color: message.role === 'assistant' ? 'C7D8D2' : 'E6D4C6', size: 4 },
          },
        }),
      );
      children.push(...markdownToDocxBlocks(message.text, spacing));
    });
  });

  return new Document({
    creator: 'Gemini Export Studio',
    styles: {
      default: {
        document: {
          run: { font: 'Aptos', size: 22, color: '1F2933' },
          paragraph: { spacing: { line: 300 } },
        },
      },
      paragraphStyles: [
        {
          id: 'Title',
          name: 'Title',
          basedOn: 'Normal',
          next: 'Normal',
          run: { font: 'Aptos Display', size: 42, bold: true, color: accent },
          paragraph: { spacing: { before: 0, after: 180 } },
        },
        {
          id: 'Heading1',
          name: 'Heading 1',
          basedOn: 'Normal',
          next: 'Normal',
          run: { font: 'Aptos Display', size: 30, bold: true, color: accent },
          paragraph: { spacing: { before: 320, after: 120 } },
        },
        {
          id: 'Heading2',
          name: 'Heading 2',
          basedOn: 'Normal',
          next: 'Normal',
          run: { font: 'Aptos Display', size: 25, bold: true, color: '293B46' },
          paragraph: { spacing: { before: 240, after: 100 } },
        },
      ],
    },
    sections: [
      {
        properties: {
          page: {
            margin: { top: 1080, right: 1080, bottom: 1080, left: 1080 },
          },
        },
        children,
      },
    ],
  });
}

function markdownToDocxBlocks(text: string, spacing: number): (Paragraph | Table)[] {
  const blocks: (Paragraph | Table)[] = [];
  const lines = text.replace(/\r\n/g, '\n').split('\n');
  let paragraphBuffer: string[] = [];

  const flushParagraph = () => {
    const content = paragraphBuffer.join(' ').trim();
    paragraphBuffer = [];
    if (content) blocks.push(...paragraphWithMath(content, spacing));
  };

  for (let index = 0; index < lines.length; index += 1) {
    const line = lines[index];
    if (!line.trim()) {
      flushParagraph();
      continue;
    }
    const fence = line.match(/^```(\w+)?/);
    if (fence) {
      flushParagraph();
      const code: string[] = [];
      index += 1;
      while (index < lines.length && !lines[index].startsWith('```')) {
        code.push(lines[index]);
        index += 1;
      }
      blocks.push(
        new Paragraph({
          children: [new TextRun({ text: code.join('\n'), font: 'Menlo', size: 18, color: '1F2937' })],
          shading: { type: ShadingType.SOLID, color: 'F3F6F8' },
          border: { left: { style: BorderStyle.SINGLE, color: '9CB5B2', size: 12 } },
          spacing: { before: 80, after: spacing },
        }),
      );
      continue;
    }
    const table = collectTable(lines, index);
    if (table) {
      flushParagraph();
      blocks.push(buildTable(table.rows));
      index = table.endIndex;
      continue;
    }
    const heading = line.match(/^(#{1,3})\s+(.+)$/);
    if (heading) {
      flushParagraph();
      blocks.push(
        new Paragraph({
          children: inlineRuns(cleanInline(heading[2])),
          heading: heading[1].length === 1 ? HeadingLevel.HEADING_2 : HeadingLevel.HEADING_3,
          spacing: { before: 180, after: 80 },
        }),
      );
      continue;
    }
    const list = line.match(/^\s*(?:[-*+]|(\d+)[.)])\s+(.+)$/);
    if (list) {
      flushParagraph();
      blocks.push(
        new Paragraph({
          children: [new TextRun({ text: list[1] ? `${list[1]}. ` : '• ', bold: true }), ...inlineRuns(cleanInline(list[2]))],
          indent: { left: 360, hanging: 180 },
          spacing: { after: 80 },
        }),
      );
      continue;
    }
    const quote = line.match(/^>\s?(.+)$/);
    if (quote) {
      flushParagraph();
      blocks.push(
        new Paragraph({
          children: inlineRuns(cleanInline(quote[1])),
          indent: { left: 300 },
          border: { left: { style: BorderStyle.SINGLE, color: 'B9C8C5', size: 12 } },
          shading: { type: ShadingType.SOLID, color: 'F7FAFA' },
          spacing: { before: 80, after: spacing },
        }),
      );
      continue;
    }
    paragraphBuffer.push(line);
  }
  flushParagraph();
  return blocks;
}

function paragraphWithMath(content: string, spacing: number): Paragraph[] {
  const result: Paragraph[] = [];
  const displayRegex = /\$\$([\s\S]+?)\$\$/g;
  let lastIndex = 0;
  for (const match of content.matchAll(displayRegex)) {
    const before = content.slice(lastIndex, match.index).trim();
    if (before) {
      result.push(new Paragraph({ children: inlineRuns(cleanInline(before)), spacing: { after: spacing } }));
    }
    result.push(
      new Paragraph({
        children: [new TextRun({ text: cleanFormula(match[1]), font: 'Cambria Math', size: 26, color: '173B46' })],
        alignment: AlignmentType.CENTER,
        spacing: { before: 100, after: spacing },
        shading: { type: ShadingType.SOLID, color: 'F5FAF9' },
        border: { top: { style: BorderStyle.SINGLE, color: 'D7E5E2', size: 4 }, bottom: { style: BorderStyle.SINGLE, color: 'D7E5E2', size: 4 } },
      }),
    );
    lastIndex = (match.index ?? 0) + match[0].length;
  }
  const rest = content.slice(lastIndex).trim();
  if (rest) result.push(new Paragraph({ children: inlineRuns(cleanInline(rest)), spacing: { after: spacing } }));
  return result;
}

function inlineRuns(text: string): TextRun[] {
  const runs: TextRun[] = [];
  const regex = /(`[^`]+`|\*\*[^*]+\*\*|\*[^*]+\*|\$[^$\n]+\$|\[[^\]]+\]\([^)]+\))/g;
  let lastIndex = 0;
  for (const match of text.matchAll(regex)) {
    const start = match.index ?? 0;
    if (start > lastIndex) runs.push(new TextRun({ text: text.slice(lastIndex, start) }));
    const token = match[0];
    if (token.startsWith('**')) runs.push(new TextRun({ text: token.slice(2, -2), bold: true }));
    else if (token.startsWith('*')) runs.push(new TextRun({ text: token.slice(1, -1), italics: true }));
    else if (token.startsWith('`')) runs.push(new TextRun({ text: token.slice(1, -1), font: 'Menlo', size: 19, color: '334155' }));
    else if (token.startsWith('$')) runs.push(new TextRun({ text: cleanFormula(token.slice(1, -1)), font: 'Cambria Math', color: '173B46' }));
    else {
      const link = token.match(/^\[([^\]]+)\]\(([^)]+)\)$/);
      runs.push(new TextRun({ text: link?.[1] ?? token, color: '245B5B', underline: {} }));
      if (link?.[2]) runs.push(new TextRun({ text: ` (${link[2]})`, color: '64748B', size: 18 }));
    }
    lastIndex = start + token.length;
  }
  if (lastIndex < text.length) runs.push(new TextRun({ text: text.slice(lastIndex) }));
  return runs.length ? runs : [new TextRun({ text })];
}

function collectTable(lines: string[], startIndex: number): { rows: string[][]; endIndex: number } | null {
  if (!lines[startIndex].includes('|') || !lines[startIndex + 1]?.match(/^\s*\|?\s*:?-{3,}:?\s*(\|\s*:?-{3,}:?\s*)+\|?\s*$/)) return null;
  const rows: string[][] = [];
  let index = startIndex;
  while (index < lines.length && lines[index].includes('|')) {
    if (!lines[index].match(/^\s*\|?\s*:?-{3,}:?\s*(\|\s*:?-{3,}:?\s*)+\|?\s*$/)) {
      rows.push(lines[index].split('|').map((cell) => cleanInline(cell.trim())).filter((cell, cellIndex, list) => cell || (cellIndex > 0 && cellIndex < list.length - 1)));
    }
    index += 1;
  }
  return { rows, endIndex: index - 1 };
}

function buildTable(rows: string[][]) {
  const columnCount = Math.max(...rows.map((row) => row.length), 1);
  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows: rows.map(
      (row, rowIndex) =>
        new TableRow({
          tableHeader: rowIndex === 0,
          children: Array.from({ length: columnCount }, (_, index) => {
            const text = row[index] ?? '';
            return new TableCell({
              shading: rowIndex === 0 ? { type: ShadingType.SOLID, color: 'E8F1EF' } : undefined,
              margins: { top: 120, bottom: 120, left: 140, right: 140 },
              children: [
                new Paragraph({
                  children: inlineRuns(text),
                  spacing: { after: 0 },
                }),
              ],
            });
          }),
        }),
    ),
  });
}

function cleanInline(text: string) {
  return text.replace(/!\[([^\]]*)\]\([^)]+\)/g, '$1').replace(/^[-*_]{3,}$/, '').trim();
}

function cleanFormula(text: string) {
  return text.replace(/\\frac\{([^{}]+)\}\{([^{}]+)\}/g, '($1)/($2)').replace(/\\pm/g, '±').replace(/\\Delta/g, 'Δ').replace(/\\sqrt\{([^{}]+)\}/g, '√($1)').trim();
}

function normalizeRole(role?: string | null): Role {
  const value = role?.toLowerCase() ?? '';
  if (['user', 'human', '你', '用户'].some((item) => value.includes(item))) return 'user';
  if (['assistant', 'model', 'gemini', 'bard', 'bot', 'ai', '模型'].some((item) => value.includes(item))) return 'assistant';
  if (['system', '系统'].some((item) => value.includes(item))) return 'system';
  return 'unknown';
}

function inferTitle(text: string) {
  return (
    text
      .replace(/^#+\s*/, '')
      .split('\n')
      .find((line) => line.trim().length > 4)
      ?.trim()
      .slice(0, 42) || '未命名会话'
  );
}

function asString(value: unknown): string | null {
  if (typeof value === 'string' && value.trim()) return value.trim();
  if (typeof value === 'number') return String(value);
  return null;
}

function makeId(...parts: string[]) {
  return parts.join('-').replace(/[^\w\u4e00-\u9fa5-]+/g, '-').slice(0, 120);
}

function dedupeConversations(items: Conversation[]) {
  const seen = new Set<string>();
  return items.filter((item) => {
    const signature = `${item.title}-${item.messages.length}-${item.messages[0]?.text.slice(0, 80)}`;
    if (seen.has(signature)) return false;
    seen.add(signature);
    return item.messages.length > 0;
  });
}

export default App;
