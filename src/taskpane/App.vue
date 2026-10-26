<script setup lang="ts">
import type { AppLocale } from '@/i18n'
import { de, en } from '@nuxt/ui/locale'
import { watchEffect } from 'vue'
import { useI18n } from 'vue-i18n'
import { useMailItem } from '@/composables/useMailItem'
import { SUPPORTED_LOCALES, useAppLocale } from '@/i18n'

const { t } = useI18n()
const locale = useAppLocale()
const item = useMailItem()

// Translations for Nuxt UI's own component texts, such as close buttons.
const uiLocales = { en, de } satisfies Record<AppLocale, typeof en>
const languages = SUPPORTED_LOCALES.map(code => ({ label: uiLocales[code].name, value: code }))

watchEffect(() => {
  document.documentElement.lang = locale.value
})
</script>

<template>
  <UApp :locale="uiLocales[locale]">
    <main class="space-y-4 p-4">
      <header class="flex items-start justify-between gap-2">
        <h1 class="text-lg font-semibold text-highlighted">
          {{ t('taskpane.title') }}
        </h1>
        <USelect
          v-model="locale"
          :items="languages"
          :aria-label="t('taskpane.language')"
          icon="i-lucide-languages"
          size="sm"
          class="w-32 shrink-0"
          data-test="language"
        />
      </header>

      <UCard v-if="item">
        <dl class="space-y-3 text-sm">
          <div>
            <dt class="font-medium text-muted">
              {{ t('taskpane.subject') }}
            </dt>
            <dd data-test="subject" class="wrap-anywhere">
              {{ item.subject || t('taskpane.noSubject') }}
            </dd>
          </div>
          <div>
            <dt class="font-medium text-muted">
              {{ t('taskpane.from') }}
            </dt>
            <dd data-test="from" class="wrap-anywhere">
              {{ item.from }}
            </dd>
          </div>
        </dl>
      </UCard>

      <UAlert
        v-else
        data-test="empty"
        icon="i-lucide-mail"
        color="neutral"
        variant="subtle"
        :title="t('taskpane.noMessage.title')"
        :description="t('taskpane.noMessage.description')"
      />
    </main>
  </UApp>
</template>
