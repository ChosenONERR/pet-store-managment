const PROXY = 'http://localhost:3456';
const PRODUCT_URL = 'https://shangoue.meituan.com/#/reuse/sc/product/page/list/product/list';

const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

async function proxy(path, options = {}) {
  const response = await fetch(`${PROXY}${path}`, options);
  const body = await response.text();
  if (!response.ok) throw new Error(`CDP proxy request failed: ${response.status} ${body}`);
  const data = JSON.parse(body);
  if (data.error) throw new Error(data.error);
  return data;
}

async function evaluate(targetId, expression) {
  const result = await proxy(`/eval?target=${encodeURIComponent(targetId)}`, {
    method: 'POST',
    headers: { 'content-type': 'text/plain; charset=utf-8' },
    body: expression,
  });
  return result.value;
}

async function waitFor(targetId, expression, timeoutMs = 90000) {
  const started = Date.now();
  while (Date.now() - started < timeoutMs) {
    if (await evaluate(targetId, expression)) return;
    await sleep(800);
  }
  throw new Error('Timed out waiting for Meituan product page');
}

async function trustedClick(targetId, selector, text) {
  const markerId = `petstar-diagnostic-${Date.now()}`;
  const prepared = await evaluate(targetId, `(() => {
    const frame = document.querySelector('#hashframe');
    const root = frame?.contentDocument;
    const element = [...(root?.querySelectorAll(${JSON.stringify(selector)}) || [])]
      .find((node) => {
        const rect = node.getBoundingClientRect();
        return node.offsetParent !== null && node.innerText.trim() === ${JSON.stringify(text)}
          && rect.width > 0 && rect.height > 0 && rect.right > 0 && rect.bottom > 0
          && rect.left < root.defaultView.innerWidth && rect.top < root.defaultView.innerHeight;
      });
    if (!frame || !root || !element) return false;
    element.scrollIntoView({ block: 'center' });
    const frameRect = frame.getBoundingClientRect();
    const rect = element.getBoundingClientRect();
    const marker = document.createElement('div');
    marker.id = ${JSON.stringify(markerId)};
    Object.assign(marker.style, { position: 'fixed', left: (frameRect.left + rect.left + rect.width / 2 - 2) + 'px',
      top: (frameRect.top + rect.top + rect.height / 2 - 2) + 'px', width: '4px', height: '4px',
      pointerEvents: 'auto', background: 'transparent', zIndex: '2147483647' });
    marker.addEventListener('click', () => element.click(), { once: true });
    document.body.appendChild(marker);
    return { x: frameRect.left + rect.left + rect.width / 2, y: frameRect.top + rect.top + rect.height / 2,
      childX: rect.left + rect.width / 2, childY: rect.top + rect.height / 2,
      childHit: root.elementFromPoint(rect.left + rect.width / 2, rect.top + rect.height / 2)?.outerHTML?.slice(0, 300) || '',
      topHit: document.elementFromPoint(frameRect.left + rect.left + rect.width / 2, frameRect.top + rect.top + rect.height / 2)?.outerHTML?.slice(0, 300) || '' };
  })()`);
  if (!prepared) throw new Error(`Could not prepare trusted click for ${text}`);
  let clickResult;
  try {
    clickResult = await proxy(`/clickAt?target=${encodeURIComponent(targetId)}`, {
      method: 'POST', headers: { 'content-type': 'text/plain; charset=utf-8' }, body: `#${markerId}`,
    });
  } finally {
    await evaluate(targetId, `document.getElementById(${JSON.stringify(markerId)})?.remove()`).catch(() => {});
  }
  return { prepared, clickResult };
}

async function main() {
  const opened = await proxy(`/new?url=${encodeURIComponent(PRODUCT_URL)}`);
  const targetId = opened.targetId;
  try {
    await waitFor(targetId, `(() => {
      const root = document.querySelector('#hashframe')?.contentDocument;
      return [...(root?.querySelectorAll('a') || [])].some((node) => node.innerText.trim() === '下载全部商品');
    })()`);
    const linkState = await evaluate(targetId, `(() => {
      const root = document.querySelector('#hashframe').contentDocument;
      const link = [...root.querySelectorAll('a')].find((node) => node.innerText.trim() === '下载全部商品');
      const rect = link.getBoundingClientRect();
      const reactPropKey = Object.keys(link).find((key) => key.startsWith('__reactProps'));
      return { outerHTML: link.outerHTML, reactPropKey: reactPropKey || '',
        reactPropNames: reactPropKey ? Object.keys(link[reactPropKey] || {}) : [],
        hasOnClick: Boolean(reactPropKey && link[reactPropKey]?.onClick),
        onClickSource: reactPropKey && link[reactPropKey]?.onClick ? String(link[reactPropKey].onClick).slice(0, 500) : '',
        rect: { left: rect.left, top: rect.top, width: rect.width, height: rect.height } };
    })()`);
    const trustedClickState = await trustedClick(targetId, 'a', '下载全部商品');
    const directHandlerState = await evaluate(targetId, `(() => {
      const root = document.querySelector('#hashframe').contentDocument;
      const link = [...root.querySelectorAll('a')].find((node) => node.innerText.trim() === '下载全部商品');
      const key = Object.keys(link).find((name) => name.startsWith('__reactProps'));
      const handler = key && link[key]?.onClick;
      if (typeof handler !== 'function') return { called: false };
      handler({ currentTarget: link, target: link, preventDefault() {}, stopPropagation() {},
        nativeEvent: { isTrusted: true }, isDefaultPrevented: () => false, isPropagationStopped: () => false });
      return { called: true };
    })()`);
    await sleep(5000);
    const state = await evaluate(targetId, `(() => {
      const frame = document.querySelector('#hashframe');
      const root = frame?.contentDocument;
      const frameRect = frame?.getBoundingClientRect();
      return {
        pageUrl: location.href,
        viewport: { innerWidth, innerHeight, devicePixelRatio, visualScale: visualViewport?.scale || 1 },
        trustedClickState: ${JSON.stringify(trustedClickState)},
        directHandlerState: ${JSON.stringify(directHandlerState)},
        linkState: ${JSON.stringify(linkState)},
        downloadLinks: [...(root?.querySelectorAll('a') || [])]
          .filter((node) => node.innerText.trim() === '下载全部商品')
          .map((node) => { const rect = node.getBoundingClientRect(); return { visible: node.offsetParent !== null,
            inViewport: rect.width > 0 && rect.height > 0 && rect.right > 0 && rect.bottom > 0
              && rect.left < root.defaultView.innerWidth && rect.top < root.defaultView.innerHeight,
            rect: { left: rect.left, top: rect.top, width: rect.width, height: rect.height } }; }),
        frameRect: frameRect ? { left: frameRect.left, top: frameRect.top, width: frameRect.width, height: frameRect.height } : null,
        modals: [...(root?.querySelectorAll('.roo-modal') || [])]
          .filter((node) => node.offsetParent !== null)
          .map((modal) => ({
            text: modal.innerText.trim().slice(0, 2000),
            buttons: [...modal.querySelectorAll('button')].map((button) => {
              const rect = button.getBoundingClientRect();
              return { text: button.innerText.trim(), visible: button.offsetParent !== null, disabled: button.disabled,
                rect: { left: rect.left, top: rect.top, width: rect.width, height: rect.height } };
            }),
          })),
        visibleDialogs: [...(root?.querySelectorAll('[role="dialog"]') || [])]
          .filter((node) => node.offsetParent !== null)
          .map((node) => node.innerText.trim().slice(0, 1000)),
      };
    })()`);
    const screenshotPath = `${process.cwd()}\\work\\meituan_inventory_diagnostic.png`;
    await proxy(`/screenshot?target=${encodeURIComponent(targetId)}&file=${encodeURIComponent(screenshotPath)}`);
    state.screenshotPath = screenshotPath;
    process.stdout.write(JSON.stringify(state, null, 2));
  } finally {
    await proxy(`/close?target=${encodeURIComponent(targetId)}`).catch(() => {});
  }
}

main().catch((error) => {
  process.stderr.write(`${error.stack || error.message}\n`);
  process.exit(1);
});
