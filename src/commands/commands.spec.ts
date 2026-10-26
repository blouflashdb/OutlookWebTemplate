import { beforeEach, describe, expect, it, vi } from 'vitest'
import { stubOffice } from '@/test-utils/office'

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
    const replaceAsync = vi.fn((_key: string, _message: unknown, callback: () => void) => callback())
    stubOffice({
      subject: 'Quarterly report',
      from: { displayName: 'Ada Lovelace', emailAddress: 'ada@contoso.com' },
      notificationMessages: { replaceAsync },
    })
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
})
