import type { Ref } from 'vue'
import { onBeforeUnmount, onMounted, shallowRef } from 'vue'

export interface MailItemSummary {
  subject: string
  from: string
}

/** Reads the message currently selected in Outlook, if any. */
export function readMailItem(): MailItemSummary | undefined {
  // The manifest only activates the add-in on messages in read mode.
  const item = Office.context.mailbox.item as Office.MessageRead | null | undefined
  if (!item)
    return undefined

  return {
    subject: item.subject,
    from: item.from.displayName || item.from.emailAddress,
  }
}

/**
 * The selected message, kept up to date while the task pane is pinned and the
 * user moves between messages.
 */
export function useMailItem(): Readonly<Ref<MailItemSummary | undefined>> {
  const item = shallowRef(readMailItem())

  onMounted(() => {
    Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, () => {
      item.value = readMailItem()
    })
  })

  onBeforeUnmount(() => {
    Office.context.mailbox.removeHandlerAsync(Office.EventType.ItemChanged)
  })

  return item
}
