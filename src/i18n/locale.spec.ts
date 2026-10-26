import { describe, expect, it } from 'vitest'
import { stubOffice } from '@/test-utils/office'
import { detectLocale, matchLocale, messages, saveLocale } from './locale'

/** Dot-separated paths of all messages, e.g. `taskpane.title`. */
function messageKeys(value: object, prefix = ''): string[] {
  return Object.entries(value).flatMap(([key, child]) =>
    typeof child === 'object' && child !== null
      ? messageKeys(child as object, `${prefix}${key}.`)
      : [`${prefix}${key}`],
  )
}

describe('matchLocale', () => {
  it.each([
    ['de-DE', 'de'],
    ['de-AT', 'de'],
    ['EN-us', 'en'],
    ['en', 'en'],
  ])('maps %s to %s', (tag, expected) => {
    expect(matchLocale(tag)).toBe(expected)
  })

  it('returns undefined for languages without translations', () => {
    expect(matchLocale('fr-FR')).toBeUndefined()
  })
})

describe('detectLocale', () => {
  it('follows Outlook\'s display language', () => {
    stubOffice({ displayLanguage: 'de-CH' })

    expect(detectLocale()).toBe('de')
  })

  it('falls back to English for languages without translations', () => {
    stubOffice({ displayLanguage: 'fr-FR' })

    expect(detectLocale()).toBe('en')
  })

  it('prefers the language the user picked in the add-in', () => {
    stubOffice({
      displayLanguage: 'en-US',
      settings: { locale: 'de' },
    })

    expect(detectLocale()).toBe('de')
  })

  it('ignores a saved language that is no longer supported', () => {
    stubOffice({
      displayLanguage: 'de-DE',
      settings: { locale: 'xx' },
    })

    expect(detectLocale()).toBe('de')
  })
})

describe('saveLocale', () => {
  it('stores the choice in the roaming settings', () => {
    const office = stubOffice()

    saveLocale('de')

    expect(office.roamingSettings.set).toHaveBeenCalledWith('locale', 'de')
    expect(office.roamingSettings.saveAsync).toHaveBeenCalledOnce()
    expect(detectLocale()).toBe('de')
  })
})

describe('messages', () => {
  it('has the same keys in every language', () => {
    const english = messageKeys(messages.en)

    for (const translation of Object.values(messages))
      expect(messageKeys(translation)).toEqual(english)
  })
})
