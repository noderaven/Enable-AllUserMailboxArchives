# Enable-AllUserMailboxArchives
This script connects to Exchange Online, retrieves a list of all regular user mailboxes, filters out any that already have an archive enabled, and then enables the archive for the remaining mailboxes. It specifically excludes shared, room, equipment, and other non-user mailbox types by only targeting 'UserMailbox' recipients.
