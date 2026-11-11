// 最小 XML 读取器：只服务于受信结构的 OOXML 部件，不依赖节点出现顺序之外的假设。
const NAMED_ENTITIES = {
  amp: '&', lt: '<', gt: '>', quot: '"', apos: "'",
};

export function decodeXml(text) {
  return text.replace(/&(#x?[0-9a-fA-F]+|[a-zA-Z][a-zA-Z0-9]*);/g, (whole, body) => {
    if (body[0] === '#') {
      const code = body[1] === 'x' || body[1] === 'X'
        ? parseInt(body.slice(2), 16)
        : parseInt(body.slice(1), 10);
      if (!Number.isFinite(code) || code < 0) return whole;
      try { return String.fromCodePoint(code); } catch { return whole; }
    }
    return Object.hasOwn(NAMED_ENTITIES, body) ? NAMED_ENTITIES[body] : whole;
  });
}

function localName(tag) {
  const i = tag.indexOf(':');
  return i < 0 ? tag : tag.slice(i + 1);
}

// 流式解析：onOpen(路径数组, 属性), onText(路径数组, 文本), onClose(路径数组)
export function parseXml(xml, handlers = {}) {
  const s = String(xml);
  const n = s.length;
  let i = 0;
  const stack = [];
  const onOpen = handlers.onOpen ?? (() => {});
  const onText = handlers.onText ?? (() => {});
  const onClose = handlers.onClose ?? (() => {});
  while (i < n) {
    if (s.charCodeAt(i) !== 60 /* < */) {
      const j = s.indexOf('<', i);
      const end = j < 0 ? n : j;
      const text = decodeXml(s.slice(i, end));
      if (text.trim() !== '') onText(stack, text);
      i = end;
      continue;
    }
    const mark = s.charCodeAt(i + 1);
    if (mark === 63 /* ? */ || mark === 33 /* ! */) {
      const close = s.indexOf('>', i);
      if (close < 0) throw new Error('XML 声明或注释不完整');
      i = close + 1;
      continue;
    }
    const close = s.indexOf('>', i);
    if (close < 0) throw new Error('XML 标签不完整');
    let tagText = s.slice(i + 1, close);
    i = close + 1;
    let selfClosing = false;
    if (tagText.endsWith('/')) {
      selfClosing = true;
      tagText = tagText.slice(0, -1);
    }
    if (tagText.startsWith('/')) {
      const name = localName(tagText.slice(1).trim());
      if (stack.length === 0 || stack[stack.length - 1] !== name) {
        throw new Error('XML 闭合标签与打开标签不匹配');
      }
      stack.pop();
      onClose(stack);
      continue;
    }
    const head = tagText.trimEnd();
    const sp = head.search(/\s/);
    const name = localName(sp < 0 ? head : head.slice(0, sp));
    if (!name) throw new Error('XML 标签名为空');
    const attrs = {};
    if (sp >= 0) {
      let p = sp + 1;
      const region = head;
      while (p < region.length) {
        while (p < region.length && region.charCodeAt(p) <= 32) p++;
        if (p >= region.length) break;
        const eq = region.indexOf('=', p);
        if (eq < 0) throw new Error(`XML 属性 ${name} 缺少取值`);
        const key = region.slice(p, eq).trim();
        if (!key) throw new Error(`XML 元素 ${name} 存在空属性名`);
        p = eq + 1;
        while (p < region.length && region.charCodeAt(p) <= 32) p++;
        const quote = region[p];
        if (quote !== '"' && quote !== "'") throw new Error(`XML 属性 ${key} 未加引号`);
        const endQ = region.indexOf(quote, p + 1);
        if (endQ < 0) throw new Error(`XML 属性 ${key} 引号未闭合`);
        attrs[key] = decodeXml(region.slice(p + 1, endQ));
        p = endQ + 1;
      }
    }
    stack.push(name);
    onOpen(stack, attrs);
    if (selfClosing) {
      stack.pop();
      onClose(stack);
    }
  }
  if (stack.length !== 0) throw new Error('XML 存在未闭合标签');
}
