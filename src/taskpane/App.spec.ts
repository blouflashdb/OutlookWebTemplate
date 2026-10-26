import type { MessageStub } from '@/test-utils/office'
import USelect from '@nuxt/ui/components/Select.vue'
import ui from '@nuxt/ui/vue-plugin'
import { mount } from '@vue/test-utils'
import { describe, expect, it } from 'vitest'
import { nextTick } from 'vue'
import { createAppI18n, detectLocale } from '@/i18n'
import { stubOffice } from '@/test-utils/office'
import App from './App.vue'

// Call after stubOffice(), which provides the display language and saved settings.
function mountApp() {
  return mount(App, {
    global: {
      plugins: [
        ui,
        createAppI18n(detectLocale()),
      ],
    },
  })
}

const report: MessageStub = {
  subject: 'Quarterly report',
  from: { displayName: 'Ada Lovelace', emailAddress: 'ada@contoso.com' },
}

describe('app', () => {
  it('shows the selected message', () => {
    stubOffice({ item: report })

    const wrapper = mountApp()

    expect(wrapper.get('[data-test="subject"]').text()).toBe('Quarterly report')
    expect(wrapper.get('[data-test="from"]').text()).toBe('Ada Lovelace')
  })

  it('shows a hint when no message is selected', () => {
    stubOffice()

    const wrapper = mountApp()

    expect(wrapper.get('[data-test="empty"]').text()).toContain('No message selected')
  })

  it('follows the selection while pinned', async () => {
    const office = stubOffice({ item: report })
    const wrapper = mountApp()

    office.selectItem({ ...report, subject: 'Re: Quarterly report' })
    await nextTick()

    expect(wrapper.get('[data-test="subject"]').text()).toBe('Re: Quarterly report')
  })

  it('stops listening for selection changes when unmounted', () => {
    const office = stubOffice({ item: report })
    const wrapper = mountApp()

    wrapper.unmount()

    expect(office.mailbox.removeHandlerAsync).toHaveBeenCalledWith('olkItemSelectedChanged')
  })
})

describe('app language', () => {
  it('uses Outlook\'s display language', () => {
    stubOffice({ displayLanguage: 'de-DE' })

    const wrapper = mountApp()

    expect(wrapper.get('h1').text()).toBe('Outlook-Web-Vorlage')
    expect(wrapper.get('[data-test="empty"]').text()).toContain('Keine Nachricht ausgewählt')
    expect(document.documentElement.lang).toBe('de')
  })

  it('switches language from the language menu and saves the choice', async () => {
    const office = stubOffice({ item: report, displayLanguage: 'en-US' })
    const wrapper = mountApp()

    // Emits the select's update:modelValue; opening its dropdown needs pointer events jsdom lacks.
    await wrapper.getComponent(USelect).setValue('de')

    expect(wrapper.text()).toContain('Betreff')
    expect(document.documentElement.lang).toBe('de')
    expect(office.roamingSettings.set).toHaveBeenCalledWith('locale', 'de')
    expect(office.roamingSettings.saveAsync).toHaveBeenCalledOnce()
  })
})
