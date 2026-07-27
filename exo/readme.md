
get-nested-dls-with-restrictions.ps1
I created this one out of necessity to traverse multiple DLs and ensure that a given sender can successfully send emails to it. A single or few DLs is not a problem to check manually, but when you have a DL with hundreds of other DLs a members (and those potentially having member DLs of their own), then it becomes unwieldy.  The second problem is that accept-list can also use DLs, so the need to check those (and nested DLs if any) for the sender being a member.
