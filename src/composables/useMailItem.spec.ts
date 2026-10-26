import { describe, expect, it } from 'vitest'
import { stubOffice } from '@/test-utils/office'
import { readMailItem } from './useMailItem'

describe('readMailItem', () => {
  it('returns undefined when no message is selected', () => {
    stubOffice()

    expect(readMailItem()).toBeUndefined()
  })

  it('uses the sender display name', () => {
    stubOffice({
      item: {
        subject: 'Quarterly report',
        from: { displayName: 'Ada Lovelace', emailAddress: 'ada@contoso.com' },
      },
    })

    expect(readMailItem()).toEqual({ subject: 'Quarterly report', from: 'Ada Lovelace' })
  })

  it('falls back to the sender email address', () => {
    stubOffice({
      item: {
        subject: 'Quarterly report',
        from: { displayName: '', emailAddress: 'ada@contoso.com' },
      },
    })

    expect(readMailItem()?.from).toBe('ada@contoso.com')
  })
})
