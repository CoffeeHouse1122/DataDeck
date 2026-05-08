import fs from 'node:fs/promises'
import path from 'node:path'
import type { AppPreferences, SelectedPaths } from '../../../shared/contracts'
import { DEFAULT_FOCUS_JOURNAL_ORDER, DEFAULT_FORCE_AE_STAFF, DEFAULT_SUMMARY_OVERRIDES } from '../../../shared/contracts'

const SETTINGS_FILE = 'preferences.json'
const LEGACY_MONTHLY_TEMPLATE = '月会数据.xlsx'
const LEGACY_STAFF_TEMPLATE = '人员数据.xlsx'

function migrateTemplatePaths(paths: Partial<SelectedPaths>, defaultPaths: Partial<SelectedPaths>): Partial<SelectedPaths> {
  const migrated = { ...paths }
  if (migrated.monthlyTemplate && path.basename(migrated.monthlyTemplate) === LEGACY_MONTHLY_TEMPLATE && defaultPaths.monthlyTemplate) {
    migrated.monthlyTemplate = defaultPaths.monthlyTemplate
  }
  if (migrated.staffTemplate && path.basename(migrated.staffTemplate) === LEGACY_STAFF_TEMPLATE && defaultPaths.staffTemplate) {
    migrated.staffTemplate = defaultPaths.staffTemplate
  }
  return migrated
}

export function createDefaultPreferences(defaultPaths: Partial<SelectedPaths> = {}): AppPreferences {
  return {
    paths: { ...defaultPaths },
    closeBehavior: 'tray',
    focusJournalOrder: [...DEFAULT_FOCUS_JOURNAL_ORDER],
    forceAeStaff: [...DEFAULT_FORCE_AE_STAFF],
    summaryOverrides: { ...DEFAULT_SUMMARY_OVERRIDES }
  }
}

export async function loadPreferences(userDataPath: string, defaultPaths: Partial<SelectedPaths> = {}): Promise<AppPreferences> {
  const filePath = path.join(userDataPath, SETTINGS_FILE)

  try {
    const raw = await fs.readFile(filePath, 'utf8')
    const parsed = JSON.parse(raw) as Partial<AppPreferences>
    return {
      ...createDefaultPreferences(defaultPaths),
      ...parsed,
      paths: migrateTemplatePaths({
        ...defaultPaths,
        ...(parsed.paths ?? {})
      }, defaultPaths),
      focusJournalOrder: parsed.focusJournalOrder?.length ? parsed.focusJournalOrder : [...DEFAULT_FOCUS_JOURNAL_ORDER],
      forceAeStaff: parsed.forceAeStaff?.length ? parsed.forceAeStaff : [...DEFAULT_FORCE_AE_STAFF],
      summaryOverrides: {
        ...DEFAULT_SUMMARY_OVERRIDES,
        ...(parsed.summaryOverrides ?? {})
      }
    }
  } catch {
    return createDefaultPreferences(defaultPaths)
  }
}

export async function savePreferences(userDataPath: string, preferences: AppPreferences): Promise<void> {
  const filePath = path.join(userDataPath, SETTINGS_FILE)
  await fs.mkdir(userDataPath, { recursive: true })
  await fs.writeFile(filePath, JSON.stringify(preferences, null, 2), 'utf8')
}
