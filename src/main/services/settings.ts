import fs from 'node:fs/promises'
import path from 'node:path'
import type { AppPreferences, SelectedPaths } from '../../shared/contracts'
import { DEFAULT_FOCUS_JOURNAL_ORDER, DEFAULT_FORCE_AE_STAFF, DEFAULT_SUMMARY_OVERRIDES } from '../../shared/contracts'

const SETTINGS_FILE = 'preferences.json'
const LEGACY_MONTHLY_TEMPLATE = '月会数据.xlsx'
const LEGACY_STAFF_TEMPLATE = '人员数据.xlsx'

function normalizedPath(value: string): string {
  const resolved = path.resolve(value)
  return process.platform === 'win32' ? resolved.toLowerCase() : resolved
}

function migrateTemplatePaths(paths: Partial<SelectedPaths>, retiredPaths: Partial<SelectedPaths>): Partial<SelectedPaths> {
  const migrated = { ...paths }
  for (const key of ['monthlyTemplate', 'staffTemplate', 'editorsJournals', 'pptTemplate'] as const) {
    const selected = migrated[key]
    const retired = retiredPaths[key]
    if (!selected || !retired) continue
    const obsoletePaths = [retired]
    if (key === 'monthlyTemplate') obsoletePaths.push(path.join(path.dirname(retired), LEGACY_MONTHLY_TEMPLATE))
    if (key === 'staffTemplate') obsoletePaths.push(path.join(path.dirname(retired), LEGACY_STAFF_TEMPLATE))
    // Match full paths, not just filenames: user-owned copies must remain selected.
    if (obsoletePaths.some((candidate) => normalizedPath(candidate) === normalizedPath(selected))) {
      delete migrated[key]
    }
  }
  return migrated
}

export function createDefaultPreferences(): AppPreferences {
  return {
    paths: {},
    closeBehavior: 'tray',
    focusJournalOrder: [...DEFAULT_FOCUS_JOURNAL_ORDER],
    forceAeStaff: [...DEFAULT_FORCE_AE_STAFF],
    summaryOverrides: { ...DEFAULT_SUMMARY_OVERRIDES }
  }
}

export async function loadPreferences(userDataPath: string, retiredPaths: Partial<SelectedPaths> = {}): Promise<AppPreferences> {
  const filePath = path.join(userDataPath, SETTINGS_FILE)

  try {
    const raw = await fs.readFile(filePath, 'utf8')
    const parsed = JSON.parse(raw) as Partial<AppPreferences>
    return {
      ...createDefaultPreferences(),
      ...parsed,
      paths: migrateTemplatePaths(parsed.paths ?? {}, retiredPaths),
      focusJournalOrder: [...DEFAULT_FOCUS_JOURNAL_ORDER],
      forceAeStaff: parsed.forceAeStaff?.length ? parsed.forceAeStaff : [...DEFAULT_FORCE_AE_STAFF],
      summaryOverrides: {
        ...DEFAULT_SUMMARY_OVERRIDES,
        ...(parsed.summaryOverrides ?? {})
      }
    }
  } catch {
    return createDefaultPreferences()
  }
}

export async function savePreferences(userDataPath: string, preferences: AppPreferences): Promise<void> {
  const filePath = path.join(userDataPath, SETTINGS_FILE)
  await fs.mkdir(userDataPath, { recursive: true })
  await fs.writeFile(filePath, JSON.stringify({ ...preferences, focusJournalOrder: [...DEFAULT_FOCUS_JOURNAL_ORDER] }, null, 2), 'utf8')
}
