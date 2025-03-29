Distribution List sync

Not a turnkey script, but I use these to export DL (and all relevant info) from a source Exchange tenant, to re-create the DLs onto a new target Excange tenant.

It can work thru a staged cutover too, where the target tenant does not have all the mailboxes from the source tenant, it will temporarily create mail-contacts as DL members, and will replace these with mailboxes as batches of users are cutover from source to target.
