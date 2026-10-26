import type { WritableComputedRef } from 'vue'
import type { AppLocale, MessageSchema } from './locale'
import { computed } from 'vue'
import { createI18n, useI18n } from 'vue-i18n'
import { FALLBACK_LOCALE, matchLocale, messages, saveLocale } from './locale'

export * from './locale'

// Editor completion for message keys in `t()`. vue-i18n still accepts any string.
declare module 'vue-i18n' {
  export interface DefineLocaleMessage extends MessageSchema {}
}

export function createAppI18n(locale: AppLocale) {
  return createI18n({
    legacy: false,
    locale,
    fallbackLocale: FALLBACK_LOCALE,
    messages,
  })
}

/** The active locale. Assigning to it switches the language and saves the choice. */
export function useAppLocale(): WritableComputedRef<AppLocale> {
  const { locale } = useI18n()

  return computed({
    get: () => matchLocale(locale.value) ?? FALLBACK_LOCALE,
    set: (value) => {
      locale.value = value
      saveLocale(value)
    },
  })
}
