// Locale detection and messages without vue-i18n, so the ribbon command page
// (commands.ts) can use them without bundling Vue.
import de from './messages/de.json'
import en from './messages/en.json'

/** Languages the add-in ships translations for. */
export const SUPPORTED_LOCALES = ['en', 'de'] as const
export type AppLocale = (typeof SUPPORTED_LOCALES)[number]
export const FALLBACK_LOCALE: AppLocale = 'en'

export type MessageSchema = typeof en

/** Every locale has to provide every message that English has. */
export const messages: Record<AppLocale, MessageSchema> = { en, de }

// Stored in the mailbox's roaming settings, so the choice follows the user to every device.
const SETTING_KEY = 'locale'

export function isSupportedLocale(value: unknown): value is AppLocale {
  return typeof value === 'string' && (SUPPORTED_LOCALES as readonly string[]).includes(value)
}

/** Maps a language tag such as `de-AT` to a supported locale. */
export function matchLocale(tag: string): AppLocale | undefined {
  const language = tag.split('-')[0]?.toLowerCase()
  return isSupportedLocale(language)
    ? language
    : undefined
}

/** The language picked in the add-in, else Outlook's display language, else English. */
export function detectLocale(): AppLocale {
  const saved: unknown = Office.context.roamingSettings.get(SETTING_KEY)
  if (isSupportedLocale(saved))
    return saved

  return matchLocale(Office.context.displayLanguage) ?? FALLBACK_LOCALE
}

/** Remembers the user's language choice in their mailbox. */
export function saveLocale(locale: AppLocale): void {
  const settings = Office.context.roamingSettings
  settings.set(SETTING_KEY, locale)
  settings.saveAsync((result) => {
    if (result.status === Office.AsyncResultStatus.Failed)
      console.warn('Could not save the language setting.', result.error)
  })
}
