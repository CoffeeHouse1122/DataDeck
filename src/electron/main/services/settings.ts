import fs from 'node:fs/promises'
import path from 'node:path'
import type { AppPreferences } from '../../../shared/contracts'
import { DEFAULT_FOCUS_JOURNAL_ORDER, DEFAULT_FORCE_AE_STAFF, DEFAULT_SUMMARY_OVERRIDES } from '../../../shared/contracts'

const SETTINGS_FILE = 'preferences.json'

export function createDefaultPreferences(): AppPreferences {
  return {
    paths: {},
    closeBehavior: 'tray',
    focusJournalOrder: [...DEFAULT_FOCUS_JOURNAL_ORDER],
    forceAeStaff: [...DEFAULT_FORCE_AE_STAFF],
    summaryOverrides: { ...DEFAULT_SUMMARY_OVERRIDES }
  }
}

export async function loadPreferences(userDataPath: string): Promise<AppPreferences> {
  const filePath = path.join(userDataPath, SETTINGS_FILE)

  try {
    const raw = await fs.readFile(filePath, 'utf8')
    const parsed = JSON.parse(raw) as Partial<AppPreferences>
    return {
      ...createDefaultPreferences(),
      ...parsed,
      paths: parsed.paths ?? {},
      focusJournalOrder: parsed.focusJournalOrder?.length ? parsed.focusJournalOrder : [...DEFAULT_FOCUS_JOURNAL_ORDER],
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
  await fs.writeFile(filePath, JSON.stringify(preferences, null, 2), 'utf8')
}
