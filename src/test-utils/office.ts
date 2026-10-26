import { vi } from 'vitest'

export interface MessageStub {
  subject: string
  from: { displayName: string, emailAddress: string }
  notificationMessages?: { replaceAsync: (...args: any[]) => void }
}

export interface OfficeStubOptions {
  /** The selected message; none when omitted. */
  item?: MessageStub
  /** Outlook's UI language. */
  displayLanguage?: string
  /** Roaming settings saved by an earlier session. */
  settings?: Record<string, unknown>
}

/**
 * Installs a minimal `Office` global, since Office.js only exists inside an
 * Office host. Stubs just the APIs this add-in uses; extend it as you use more.
 */
export function stubOffice({ item, displayLanguage = 'en-US', settings = {} }: OfficeStubOptions = {}) {
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
  const stored = new Map(Object.entries(settings))
  const roamingSettings = {
    get: vi.fn((name: string) => stored.get(name) ?? null),
    set: vi.fn((name: string, value: unknown) => {
      stored.set(name, value)
    }),
    remove: vi.fn((name: string) => {
      stored.delete(name)
    }),
    saveAsync: vi.fn((callback?: (result: { status: string }) => void) => {
      callback?.({ status: 'succeeded' })
    }),
  }
  const actions = {
    associate: vi.fn(),
  }

  vi.stubGlobal('Office', {
    context: { mailbox, displayLanguage, roamingSettings },
    actions,
    onReady: vi.fn(async (callback?: () => void) => callback?.()),
    AsyncResultStatus: { Succeeded: 'succeeded', Failed: 'failed' },
    EventType: { ItemChanged: 'olkItemSelectedChanged' },
    MailboxEnums: {
      ItemNotificationMessageType: { InformationalMessage: 'informationalMessage' },
    },
  })

  return {
    mailbox,
    roamingSettings,
    actions,
    /** Simulates the user selecting another message while the task pane is pinned. */
    selectItem(next?: MessageStub) {
      mailbox.item = next
      handlers.forEach(handler => handler())
    },
  }
}
