import assert from 'node:assert/strict'
import { test } from 'node:test'
import { resolve } from 'node:path'
import { createServer } from 'vite'
import { chromium } from 'playwright'

test('only the registered game frame at the document origin may update a score', { timeout: 30000 }, async () => {
  const root = resolve(import.meta.dirname, '../packages/snake')
  const server = await createServer({ root, configFile: resolve(root, 'vite.config.ts'), server: { host: '127.0.0.1', port: 0 } })
  let browser
  try {
    await server.listen()
    const address = server.httpServer.address()
    assert(address && typeof address === 'object')
    browser = await chromium.launch()
    const page = await browser.newPage()
    await page.route('**/favicon.ico', route => route.fulfill({ status: 204 }))
    await page.goto(`http://127.0.0.1:${address.port}/snake/`)
    await page.waitForFunction(() => !!window.editor)
    const frame = page.frames().find(frame => frame.parentFrame())
    assert(frame)
    const gameId = await frame.evaluate(() => CONFIG.gameId)
    const score = () => page.evaluate(() => window.editor.command.getValue().data.main
      .find(element => element.control)?.control.value.map(element => element.value).join(''))
    const before = await score()
    // A real same-origin postMessage from the parent has an unauthorized source.
    await page.evaluate(async gameId => {
      await new Promise(resolve => {
        window.addEventListener('message', () => resolve(), { once: true })
        window.postMessage({ type: 'canvas-editor-snake:score', gameId, score: 999, duration: 999 }, location.origin)
      })
    }, gameId)
    assert.equal(await score(), before)
    // The source alone is insufficient: a different origin must also be rejected.
    await page.evaluate(gameId => window.dispatchEvent(new MessageEvent('message', {
      origin: 'https://example.com', source: document.querySelector('iframe').contentWindow,
      data: { type: 'canvas-editor-snake:score', gameId, score: 888, duration: 888 }
    })), gameId)
    assert.equal(await score(), before)
    // A sibling game cannot update the first game's score even at the same origin.
    await page.getByRole('button', { name: '插入贪吃蛇' }).click()
    await page.waitForFunction(() => document.querySelectorAll('iframe').length === 2)
    const sibling = page.frames().filter(frame => frame.parentFrame()).at(-1)
    await sibling.evaluate(gameId => window.top.postMessage({ type: 'canvas-editor-snake:score', gameId, score: 777, duration: 777 }, '*'), gameId)
    await page.evaluate(() => new Promise(resolve => setTimeout(resolve, 0)))
    assert.equal(await score(), before)
    const current = page.frames().find(frame => frame.parentFrame())
    await current.evaluate(gameId => window.top.postMessage({ type: 'canvas-editor-snake:score', gameId, score: 42, duration: 7 }, '*'), gameId)
    await page.waitForFunction(() => window.editor.command.getValue().data.main
      .find(element => element.control)?.control.value.some(element => element.value.includes('42')))
    assert.match(await score(), /42.*7/)
  } finally {
    await browser?.close()
    await server.close()
  }
})
