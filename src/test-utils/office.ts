import { vi } from 'vitest'

export interface MessageStub {
  subject: string
  from: { displayName: string, emailAddress: string }
  notificationMessages?: { replaceAsync: (...args: any[]) => void }
}

/**
 * Installs a minimal `Office` global, since Office.js only exists inside an
 * Office host. Stubs just the APIs this add-in uses; extend it as you use more.
 */
export function stubOffice(item?: MessageStub) {
  const handlers = new Set<() => void>()
  const mailbox = {
    item,
    addHandlerAsync: vi.fn((_type: string, handler: () => void) => {
      handlers.add(handler)
    }),
    removeHandlerAsync: vi.fn(() => {
      handlers.clear()
    }),
  }
  const actions = {
    associate: vi.fn(),
  }

  vi.stubGlobal('Office', {
    context: { mailbox },
    actions,
    onReady: vi.fn(async (callback?: () => void) => callback?.()),
    EventType: { ItemChanged: 'olkItemSelectedChanged' },
    MailboxEnums: {
      ItemNotificationMessageType: { InformationalMessage: 'informationalMessage' },
    },
  })

  return {
    mailbox,
    actions,
    /** Simulates the user selecting another message while the task pane is pinned. */
    selectItem(next?: MessageStub) {
      mailbox.item = next
      handlers.forEach(handler => handler())
    },
  }
}
