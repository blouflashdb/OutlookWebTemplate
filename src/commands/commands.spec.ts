import type { OfficeStubOptions } from '@/test-utils/office'
import { beforeEach, describe, expect, it, vi } from 'vitest'
import { stubOffice } from '@/test-utils/office'

function stubOfficeWithMessage(options: Omit<OfficeStubOptions, 'item'> = {}) {
  const replaceAsync = vi.fn((_key: string, _message: unknown, callback: () => void) => callback())
  stubOffice({
    ...options,
    item: {
      subject: 'Quarterly report',
      from: { displayName: 'Ada Lovelace', emailAddress: 'ada@contoso.com' },
      notificationMessages: { replaceAsync },
    },
  })
  return replaceAsync
}

describe('action', () => {
  beforeEach(() => {
    // commands.ts registers its handlers on import, so load it fresh per test.
    vi.resetModules()
  })

  it('is registered with Office', async () => {
    const office = stubOffice()

    const { action } = await import('./commands')

    expect(office.actions.associate).toHaveBeenCalledWith('action', action)
  })

  it('adds a notification to the message and completes the event', async () => {
    const replaceAsync = stubOfficeWithMessage()
    const { action } = await import('./commands')
    const event = {
      completed: vi.fn(),
    }

    action(event as unknown as Office.AddinCommands.Event)

    expect(replaceAsync).toHaveBeenCalledWith(
      'action',
      expect.objectContaining({ message: 'Performed action.' }),
      expect.any(Function),
    )
    expect(event.completed).toHaveBeenCalledOnce()
  })

  it('writes the notification in the user\'s language', async () => {
    const replaceAsync = stubOfficeWithMessage({ displayLanguage: 'de-DE' })
    const { action } = await import('./commands')

    action({
      completed: vi.fn(),
    } as unknown as Office.AddinCommands.Event)

    expect(replaceAsync).toHaveBeenCalledWith(
      'action',
      expect.objectContaining({ message: 'Aktion ausgeführt.' }),
      expect.any(Function),
    )
  })
})
