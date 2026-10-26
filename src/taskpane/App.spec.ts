import type { MessageStub } from '@/test-utils/office'
import ui from '@nuxt/ui/vue-plugin'
import { mount } from '@vue/test-utils'
import { describe, expect, it } from 'vitest'
import { nextTick } from 'vue'
import { stubOffice } from '@/test-utils/office'
import App from './App.vue'

function mountApp() {
  return mount(App, {
    global: {
      plugins: [ui],
    },
  })
}

const report: MessageStub = {
  subject: 'Quarterly report',
  from: { displayName: 'Ada Lovelace', emailAddress: 'ada@contoso.com' },
}

describe('app', () => {
  it('shows the selected message', () => {
    stubOffice(report)

    const wrapper = mountApp()

    expect(wrapper.get('[data-test="subject"]').text()).toBe('Quarterly report')
    expect(wrapper.get('[data-test="from"]').text()).toBe('Ada Lovelace')
  })

  it('shows a hint when no message is selected', () => {
    stubOffice()

    const wrapper = mountApp()

    expect(wrapper.find('[data-test="empty"]').exists()).toBe(true)
  })

  it('follows the selection while pinned', async () => {
    const office = stubOffice(report)
    const wrapper = mountApp()

    office.selectItem({ ...report, subject: 'Re: Quarterly report' })
    await nextTick()

    expect(wrapper.get('[data-test="subject"]').text()).toBe('Re: Quarterly report')
  })

  it('stops listening for selection changes when unmounted', () => {
    const office = stubOffice(report)
    const wrapper = mountApp()

    wrapper.unmount()

    expect(office.mailbox.removeHandlerAsync).toHaveBeenCalledWith('olkItemSelectedChanged')
  })
})
