import { openPetstarMerchantTab, releasePetstarMerchantTab } from './meituan_browser_context.mjs';

const PROXY = 'http://localhost:3456';
const PRODUCT_URL = 'https://shangoue.meituan.com/#/reuse/sc/product/page/list/product/list';
const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

async function proxy(requestPath, options = {}) {
  const response = await fetch(`${PROXY}${requestPath}`, {
    ...options,
    signal: options.signal || AbortSignal.timeout(30000),
  });
  const body = await response.text();
  if (!response.ok) throw new Error(`CDP proxy request failed: ${response.status} ${body}`);
  const result = JSON.parse(body);
  if (result.error) throw new Error(result.error);
  return result;
}

async function evaluate(targetId, expression) {
  const result = await proxy(`/eval?target=${encodeURIComponent(targetId)}`, {
    method: 'POST',
    headers: { 'content-type': 'text/plain; charset=utf-8' },
    body: expression,
  });
  return result.value;
}

async function waitFor(targetId, expression, timeoutMs, description) {
  const started = Date.now();
  while (Date.now() - started < timeoutMs) {
    if (await evaluate(targetId, expression)) return;
    await sleep(800);
  }
  throw new Error(`Timed out waiting for ${description}`);
}

async function main() {
  const opened = await openPetstarMerchantTab(PRODUCT_URL);
  try {
    await waitFor(opened.targetId, `(() => {
      const root = document.querySelector('#hashframe')?.contentDocument;
      return [...(root?.querySelectorAll('a') || [])].some((node) => node.innerText.trim() === '下载全部商品');
    })()`, 90000, 'the product list');

    await evaluate(opened.targetId, `(() => {
      const root = document.querySelector('#hashframe')?.contentDocument;
      const link = [...(root?.querySelectorAll('a') || [])].find((node) => node.innerText.trim() === '下载全部商品');
      if (!link) return false;
      const key = Object.keys(link).find((name) => name.startsWith('__reactProps'));
      const handler = key && link[key]?.onClick;
      if (typeof handler === 'function') {
        handler({ currentTarget: link, target: link, preventDefault() {}, stopPropagation() {},
          nativeEvent: { isTrusted: true }, isDefaultPrevented: () => false, isPropagationStopped: () => false });
      } else {
        link.click();
      }
      return true;
    })()`);

    await waitFor(opened.targetId, `(() => {
      const root = document.querySelector('#hashframe')?.contentDocument;
      const modal = [...(root?.querySelectorAll('.roo-modal') || [])]
        .filter((node) => node.offsetParent !== null)
        .find((node) => /新增下载文件|下载文件|操作时间|生成状态|任务状态/.test(node.innerText || ''));
      return Boolean(modal && !/正在加载/.test(modal.innerText || '') && modal.querySelectorAll('input[type="checkbox"]').length);
    })()`, 60000, 'the export fields');

    const fields = await evaluate(opened.targetId, `(() => {
      const root = document.querySelector('#hashframe')?.contentDocument;
      const modal = [...root.querySelectorAll('.roo-modal')]
        .filter((node) => node.offsetParent !== null)
        .find((node) => /新增下载文件|下载文件|操作时间|生成状态|任务状态/.test(node.innerText || ''));
      return [...modal.querySelectorAll('input[type="checkbox"]')].map((input) => {
        const label = input.closest('label');
        const container = label || input.parentElement;
        return {
          value: input.value || '',
          checked: Boolean(input.checked),
          text: (container?.innerText || container?.textContent || '').trim().replace(/\\s+/g, ' '),
          outer: (container?.outerHTML || input.outerHTML || '').slice(0, 600),
        };
      });
    })()`);
    process.stdout.write(JSON.stringify({ ok: true, source: opened.source, fields }, null, 2));
  } finally {
    await releasePetstarMerchantTab(opened);
  }
}

main().catch((error) => {
  process.stderr.write(`${error.stack || error.message}\n`);
  process.exit(1);
});
