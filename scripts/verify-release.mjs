import assert from 'node:assert/strict'
import fs from 'node:fs/promises'
import path from 'node:path'
import { createHash } from 'node:crypto'
import { createRequire } from 'node:module'
import { fileURLToPath } from 'node:url'

// Both parsers are supplied by the locked electron-builder toolchain.
const require = createRequire(import.meta.url)
const asar = require('@electron/asar')
const yaml = require('js-yaml')
const root = fileURLToPath(new URL('../', import.meta.url))

export function verifyTag(version, tag) {
  assert.match(version, /^\d+\.\d+\.\d+$/, 'Only stable releases are supported')
  assert.equal(tag, `v${version}`, 'Tag must match package.json version')
}

export async function verifyRelease(directory, version) {
  const metadata = yaml.load(await fs.readFile(path.join(directory, 'latest.yml'), 'utf8'))
  assert.equal(metadata.version, version)
  const filename = `DataDeck-${version}-Setup.exe`
  assert.equal(metadata.files.length, 1, 'Expected one Windows x64 installer')
  const info = metadata.files[0]
  assert.equal(info.url, filename)
  const installer = await fs.readFile(path.join(directory, filename))
  assert.equal(info.size, installer.length)
  assert.equal(info.sha512, createHash('sha512').update(installer).digest('base64'))
  assert.ok((await fs.stat(path.join(directory, `${filename}.blockmap`))).size > 0)

  const resources = path.join(directory, 'win-unpacked', 'resources')
  const config = yaml.load(await fs.readFile(path.join(resources, 'app-update.yml'), 'utf8'))
  assert.equal(config.provider, 'github')
  assert.equal(config.owner, 'CoffeeHouse1122')
  assert.equal(config.repo, 'DataDeck')
  assert.equal(config.publisherName, undefined, 'Unsigned release must not require a publisher certificate')
  assert.notEqual(config.private, true)
  assert.equal(config.token, undefined, 'No credentials may be embedded')
  const archive = path.join(resources, 'app.asar')
  // A caller may have rebuilt this archive at the same path; never inspect cached headers.
  asar.uncache(archive)
  const entries = asar.listPackage(archive).map((name) => name.replace(/\\/g, '/'))
  const privatePath = /\.(xlsx?|xlsm|xlsb|pptx?|pptm|docx?|docm)$/i
  const forbiddenDirectory = /(^|\/)(private-business-data|templates)(\/|$)/i
  for (const entry of entries) {
    assert.ok(!privatePath.test(entry) && !forbiddenDirectory.test(entry) && !entry.startsWith('/docs/'), `Private resource in app.asar: ${entry}`)
  }
  async function checkLooseFiles(directory) {
    for (const entry of await fs.readdir(directory, { withFileTypes: true })) {
      assert.ok(!privatePath.test(entry.name) && !['templates', 'private-business-data'].includes(entry.name.toLowerCase()), `Private resource in resources: ${entry.name}`)
      if (entry.isDirectory()) await checkLooseFiles(path.join(directory, entry.name))
    }
  }
  await checkLooseFiles(resources)
  const packaged = JSON.parse(asar.extractFile(archive, 'package.json').toString('utf8'))
  assert.equal(packaged.version, version)
  for (const entry of ['/dist-electron/main/index.js', '/dist-electron/preload/index.mjs', '/dist-electron/renderer/index.html']) {
    assert.ok(entries.includes(entry), `Missing application entry: ${entry}`)
  }
}

if (process.argv[1] && path.resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  const { version } = JSON.parse(await fs.readFile(path.join(root, 'package.json'), 'utf8'))
  if (process.argv.includes('--tag')) verifyTag(version, process.env.GITHUB_REF_NAME)
  else await verifyRelease(path.join(root, 'release'), version)
  console.log('PASS: release verification')
}
