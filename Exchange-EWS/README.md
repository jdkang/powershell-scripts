Powershell example(s) of the [EWS API](https://www.google.com/#q=EWS+api).

Originally written for Exchange 2007 and EWS API 1.7, updating to 2.0 and newer Exchange versions should be pretty straight forward. Newer versions of exchange offer newer options as well.

Some code has been commenting out for performance reasons--like grabbing headers or more complex filter examples.

The `Grant-Impersonation.ps1` script was written for granting rights for a mailbox across CAS as per Exchange 2007 methodology. You'll still need to grant impersonate rights as per the paradigm of your Exchange version (keeping in mind least privilege). 