import assert from 'node:assert/strict'
import fs from 'node:fs/promises'
import path from 'node:path'
import { createRequire } from 'node:module'
import { fileURLToPath } from 'node:url'
import { build } from 'vite'

const root = fileURLToPath(new URL('../', import.meta.url))
const require = createRequire(import.meta.url)
const minimatch = require('minimatch')
const config = JSON.parse(await fs.readFile(path.join(root, 'package.json'), 'utf8')).build

// Exercise the packaging allow/exclude rules against private and runtime files.
assert.deepEqual(config.extraResources, [{ from: 'resources/icons', to: 'icons' }])
assert.equal(config.extraFiles, undefined)
const included = (filename) => config.files.some((pattern) => !pattern.startsWith('!') && minimatch(filename, pattern, { dot: true }))
  && !config.files.some((pattern) => pattern.startsWith('!') && minimatch(filename, pattern.slice(1), { dot: true }))
for (const filename of ['docs/private.xlsx', 'docs/private.pptx', 'private-business-data/使用说明.txt', 'private-business-data/private.xlsx', 'dist/old-template.xlsx', 'dist-electron/private.xlsx', 'dist-electron/private.pptx', 'dist-electron/private.docx', 'dist-electron/templates/screenshot.png']) {
  assert.equal(included(filename), false, `${filename} must not be packaged`)
}
for (const filename of ['dist-electron/main/index.js', 'dist-electron/preload/index.mjs', 'dist-electron/renderer/index.html', 'dist-electron/renderer/assets/font.woff2', 'package.json']) {
  assert.equal(included(filename), true, `${filename} must remain packaged`)
}
console.log('PASS: private resources excluded; application entries and assets retained')

await fs.mkdir(path.join(root, 'out'), { recursive: true })
const output = await fs.mkdtemp(path.join(root, 'out', 'private-resources-check-'))
await build({
  configFile: false, root, logLevel: 'error',
  build: {
    ssr: path.join(root, 'src/main/services/settings.ts'), outDir: output, emptyOutDir: false,
    rollupOptions: { output: { format: 'cjs', entryFileNames: 'settings.cjs' } }
  }
})
const { createDefaultPreferences, loadPreferences, savePreferences } = require(path.join(output, 'settings.cjs'))
const userData = path.join(output, 'user-data')
const retiredDirectory = path.join(output, 'old-install', 'resources', 'templates')
const retired = {
  monthlyTemplate: path.join(retiredDirectory, 'monthly-data-generated-202603.xlsx'),
  staffTemplate: path.join(retiredDirectory, 'staff-data-generated-202603.xlsx'),
  editorsJournals: path.join(retiredDirectory, 'editors-journals.xlsx'),
  pptTemplate: path.join(retiredDirectory, 'Section Health月会.pptx')
}
assert.deepEqual(createDefaultPreferences().paths, {})
assert.deepEqual((await loadPreferences(userData, retired)).paths, {})
console.log('PASS: fresh installation does not select bundled resources')

const preferences = {
  ...createDefaultPreferences(), closeBehavior: 'quit', forceAeStaff: ['Demo User'],
  paths: { ...retired, mrWorkbook: path.join(output, 'selected-mr.xlsx'), outputDir: path.join(output, 'reports') },
  summaryOverrides: { ...createDefaultPreferences().summaryOverrides, revenueWCHF: '12.34' }
}
// Leftover files from an old installation must not restore bundled defaults.
await fs.mkdir(retiredDirectory, { recursive: true })
await fs.writeFile(retired.monthlyTemplate, 'synthetic migration fixture')
await savePreferences(userData, preferences)
const migrated = await loadPreferences(userData, retired)
assert.deepEqual(migrated.paths, { mrWorkbook: preferences.paths.mrWorkbook, outputDir: preferences.paths.outputDir })
assert.equal(migrated.closeBehavior, preferences.closeBehavior)
assert.deepEqual(migrated.forceAeStaff, preferences.forceAeStaff)
assert.deepEqual(migrated.summaryOverrides, preferences.summaryOverrides)
await savePreferences(userData, migrated)
assert.deepEqual(await loadPreferences(userData, retired), migrated)

await savePreferences(userData, { ...preferences, paths: {
  monthlyTemplate: path.join(retiredDirectory, '月会数据.xlsx'),
  staffTemplate: path.join(retiredDirectory, '人员数据.xlsx')
} })
assert.deepEqual((await loadPreferences(userData, retired)).paths, {})
if (process.platform === 'win32') {
  await savePreferences(userData, { ...preferences, paths: { monthlyTemplate: retired.monthlyTemplate.toUpperCase() } })
  assert.deepEqual((await loadPreferences(userData, retired)).paths, {})
}
console.log('PASS: obsolete bundled paths cleared; other preferences preserved')

const customPaths = Object.fromEntries(Object.entries(retired).map(([key, filename]) => [key, path.join(output, 'user-selected', path.basename(filename))]))
customPaths.staffTemplate = path.join(output, 'user-selected', '人员数据.xlsx')
await savePreferences(userData, { ...preferences, paths: customPaths })
assert.deepEqual((await loadPreferences(userData, retired)).paths, customPaths)
console.log('PASS: external paths with matching current/legacy filenames preserved')
console.log(`Artifacts: ${output}`)
