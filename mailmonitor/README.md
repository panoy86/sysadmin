This requires 2 Microsoft 365 tenants.  Presumably, the "1_" sending script will use a service principal on the production tenant.
While the "2_" check-email script will use a service principal on the test/dev tenant.

This also can only work if the outbound path of an email sent from production to the test/dev tenant follows the same path as all other externally-bound emails.

Enjoy - franaur@gmail.com
