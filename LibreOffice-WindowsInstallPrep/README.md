Requires a [LibreOffice MSI](https://www.libreoffice.org/download/libreoffice-fresh/), which will get extracted by `Stage-LibreOfficeMsi.ps1`

Configures LibreOffice to save files as MS Office file formats (e.g. .docx) **by default**.

Also drops in a .BAT file which has pre-set MSI installation arguments to:

1. Set file associations for MS office types
2. Disable EULA
3. Disable checking for updates
4. Disable online updates
5. Log
