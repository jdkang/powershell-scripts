# IIS Orchestration

`bootstrap_00_iis_Server2012.ps1` - Setup IIS features and move default `inetpub`

`bootstrap_01_webdeploy.ps1` - Install WebDeploy, enable IIS backups, delegation handler, and setup a delegate user with rules

`bootstrap_10_F5NativeModule.ps1` - Install the [F5 IIS X-FORWARDED-FOR](https://devcentral.f5.com/articles/x-forwarded-for-http-module-for-iis7-source-included) module

`sitewebapp\setup_webapp-example.ps1` - Example of setting up a new IIS webapp using the `webadministration` module