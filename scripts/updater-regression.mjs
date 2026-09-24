import assert from 'node:assert/strict'
import fs from 'node:fs/promises'
import path from 'node:path'
import { EventEmitter } from 'node:events'
import { createRequire } from 'node:module'
import { fileURLToPath } from 'node:url'
import { createHash } from 'node:crypto'
import { build } from 'vite'
import vue from '@vitejs/plugin-vue'
import { createSSRApp } from 'vue'
import { renderToString } from '@vue/server-renderer'
import { verifyRelease, verifyTag } from './verify-release.mjs'

const root = fileURLToPath(new URL('../', import.meta.url))
const require = createRequire(import.meta.url)
await fs.mkdir(path.join(root, 'out'), { recursive: true })
const output = await fs.mkdtemp(path.join(root, 'out', 'updater-check-'))
await build({ configFile: false, root, logLevel: 'error', build: {
  ssr: path.join(root, 'src/main/services/updater.ts'), outDir: output, emptyOutDir: false,
  rollupOptions: { output: { format: 'cjs', entryFileNames: 'updater.cjs' } }
} })
const { UpdateController } = require(path.join(output, 'updater.cjs'))
const defer = () => { let resolve; const promise = new Promise((r) => { resolve = r }); return { promise, resolve } }
class FakeUpdater extends EventEmitter {
  checks = 0; downloads = 0; installs = 0
  checkForUpdates() { this.checks++; return this.checkResult ?? Promise.resolve() }
  downloadUpdate() { this.downloads++; return this.downloadResult ?? Promise.resolve() }
  quitAndInstall(...args) { this.installs++; this.installArgs = args; if (this.installError) throw new Error('install failure') }
}
const client = new FakeUpdater()
let quitting = false
const updates = new UpdateController(client, true, '0.1.4', () => {}, (value) => { quitting = value })
assert.equal(client.autoDownload, false)
assert.equal(client.autoInstallOnAppQuit, false)
assert.equal(client.allowPrerelease, false)
assert.equal(client.allowDowngrade, false)
const checking = defer(); client.checkResult = checking.promise
const firstCheck = updates.check(); await updates.check()
assert.equal(client.checks, 1)
client.emit('update-available', { version: '0.1.5' }); checking.resolve(); await firstCheck
assert.equal(updates.getState().phase, 'available')
const downloading = defer(); client.downloadResult = downloading.promise
const firstDownload = updates.download(); await updates.download(); await updates.check()
assert.equal(client.downloads, 1); assert.equal(client.checks, 1)
client.emit('download-progress', { percent: 37.2 })
assert.equal(updates.getState().percent, 37.2)
client.emit('update-downloaded', { version: '0.1.5' }); downloading.resolve(); await firstDownload
await updates.check(); assert.equal(client.checks, 1)
const generation = defer()
const run = updates.runGeneration(() => generation.promise)
assert.throws(() => updates.install(), /报表正在生成/)
await assert.rejects(updates.runGeneration(async () => {}), /已有报表/)
assert.equal(client.installs, 0)
generation.resolve('report'); assert.equal(await run, 'report')
await assert.rejects(updates.runGeneration(async () => { throw new Error('write failed') }), /write failed/)
assert.equal(updates.getState().generationRunning, false)
updates.install(); assert.equal(quitting, true); assert.deepEqual(client.installArgs, [false, true])
await assert.rejects(updates.runGeneration(async () => {}), /正在安装更新/)
assert.throws(() => updates.install(), /请先下载/)
client.emit('error', new Error('https://example.invalid/private-response'))
assert.equal(quitting, false); assert.equal(updates.getState().phase, 'error')
assert.ok(!updates.getState().message.includes('https'))
client.checkResult = Promise.reject(new Error('offline')); await updates.check()
assert.equal(updates.getState().phase, 'error')
client.checkResult = Promise.resolve(); const retry = updates.check()
client.emit('update-available', { version: '0.1.5' }); await retry
client.downloadResult = Promise.reject(new Error('network interrupted')); await updates.download()
assert.equal(updates.getState().phase, 'error')
client.emit('update-downloaded', { version: '0.1.5' }); client.installError = true
updates.install(); assert.equal(quitting, false); assert.equal(updates.getState().phase, 'error')
await updates.check(); client.emit('update-not-available')
assert.equal(updates.getState().phase, 'idle')
const disabledClient = new FakeUpdater()
const disabled = new UpdateController(disabledClient, false, '0.1.4', () => {}, () => {})
disabled.start(); await disabled.check(); await disabled.download(); disabled.stop()
assert.equal(disabledClient.checks, 0); assert.equal(disabledClient.downloads, 0)
assert.equal(await disabled.runGeneration(async () => 'report'), 'report')
console.log('PASS: manual updates, duplicate calls, retry, generation/install exclusion, development mode')

await build({ configFile: false, root, logLevel: 'error', plugins: [vue()], build: {
  ssr: path.join(root, 'src/renderer/components/UpdatePanel.vue'), outDir: output, emptyOutDir: false,
  rollupOptions: { output: { format: 'cjs', entryFileNames: 'update-panel.cjs' } }
} })
const UpdatePanel = require(path.join(output, 'update-panel.cjs'))
const panel = (phase, extra = {}, props = {}) => renderToString(createSSRApp(UpdatePanel, {
  state: { phase, currentVersion: '0.1.4', generationRunning: false, message: 'update status', ...extra },
  pending: false, running: false, ...props
}))
assert.match(await panel('downloaded', { generationRunning: true }), /<button[^>]*disabled[^>]*>保存并重启安装/)
assert.match(await panel('downloaded', {}, { running: true }), /<button[^>]*disabled[^>]*>保存并重启安装/)
assert.doesNotMatch(await panel('downloaded'), /<button[^>]*disabled/)
assert.match(await panel('downloading', { percent: 37 }), /<progress[^>]*value="37"/)
assert.match(await panel('disabled'), /<button[^>]*disabled/)
assert.match(await panel('error'), /重新检查更新/)
console.log('PASS: update panel progress, install guards and retry state')

// Synthetic packages exercise the release gate without touching real business files or making an installer.
const asar = require('@electron/asar'); const yaml = require('js-yaml')
const release = path.join(output, 'release'); const resources = path.join(release, 'win-unpacked', 'resources')
const source = path.join(output, 'synthetic-app')
for (const entry of ['main/index.js', 'preload/index.mjs', 'renderer/index.html']) {
  const filename = path.join(source, 'dist-electron', entry)
  await fs.mkdir(path.dirname(filename), { recursive: true }); await fs.writeFile(filename, 'synthetic')
}
await fs.writeFile(path.join(source, 'package.json'), JSON.stringify({ version: '0.1.4' }))
await fs.mkdir(resources, { recursive: true })
await asar.createPackage(source, path.join(resources, 'app.asar'))
const configuration = { provider: 'github', owner: 'CoffeeHouse1122', repo: 'DataDeck' }
await fs.writeFile(path.join(resources, 'app-update.yml'), yaml.dump(configuration))
const installer = Buffer.from('synthetic installer'); const filename = 'DataDeck-0.1.4-Setup.exe'
await fs.writeFile(path.join(release, filename), installer)
await fs.writeFile(path.join(release, `${filename}.blockmap`), 'synthetic blockmap')
await fs.writeFile(path.join(release, 'latest.yml'), yaml.dump({ version: '0.1.4', files: [{ url: filename, size: installer.length, sha512: createHash('sha512').update(installer).digest('base64') }] }))
verifyTag('0.1.4', 'v0.1.4')
assert.throws(() => verifyTag('0.1.4', 'v0.1.3'))
assert.throws(() => verifyTag('0.1.4-beta.1', 'v0.1.4-beta.1'))
await verifyRelease(release, '0.1.4')
await fs.writeFile(path.join(source, 'private.xlsx'), 'synthetic private fixture')
await asar.createPackage(source, path.join(resources, 'app.asar'))
await assert.rejects(verifyRelease(release, '0.1.4'), /Private resource/)
await fs.writeFile(path.join(release, filename), 'tampered')
await assert.rejects(verifyRelease(release, '0.1.4'))
const workflow = yaml.load(await fs.readFile(path.join(root, '.github/workflows/release.yml'), 'utf8'))
assert.equal(workflow.on.push.tags[0], 'v*')
assert.equal(workflow.permissions.contents, 'write')
const config = JSON.parse(await fs.readFile(path.join(root, 'package.json'), 'utf8'))
assert.deepEqual(config.build.publish, { ...configuration, private: false })
assert.equal(config.build.win.verifyUpdateCodeSignature, false)
assert.ok(config.scripts['dist:win'].includes('--publish never'))
console.log('PASS: release tag, artifact hash/entries/feed, private resource rejection, workflow syntax')
console.log(`Artifacts: ${output}`)
