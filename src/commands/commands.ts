/**
 * Handler for the "Perform an Action" ribbon button. Function names must match
 * <FunctionName> in manifest.xml.
 */
export function action(event: Office.AddinCommands.Event): void {
  const message: Office.NotificationMessageDetails = {
    type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
    message: 'Performed action.',
    icon: 'Icon.80x80',
    persistent: true,
  }

  Office.context.mailbox.item?.notificationMessages.replaceAsync('action', message, () => {
    // Tell Outlook the command is done; it times out the command otherwise.
    event.completed()
  })
}

void Office.onReady(() => {
  Office.actions.associate('action', action)
})
