<script setup lang="ts">
import { useMailItem } from '@/composables/useMailItem'

const item = useMailItem()
</script>

<template>
  <UApp>
    <main class="space-y-4 p-4">
      <h1 class="text-lg font-semibold text-highlighted">
        Outlook Web Template
      </h1>

      <UCard v-if="item">
        <dl class="space-y-3 text-sm">
          <div>
            <dt class="font-medium text-muted">
              Subject
            </dt>
            <dd data-test="subject" class="wrap-anywhere">
              {{ item.subject || '(no subject)' }}
            </dd>
          </div>
          <div>
            <dt class="font-medium text-muted">
              From
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
        title="No message selected"
        description="Select a message to see its details."
      />
    </main>
  </UApp>
</template>
