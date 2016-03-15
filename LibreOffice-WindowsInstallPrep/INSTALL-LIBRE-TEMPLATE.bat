Pushd "%~dp0"

REM https://wiki.documentfoundation.org/Deployment_and_Migration
REM http://wpkg.org/LibreOffice
REM http://listarchives.libreoffice.org/global/users/msg12888.html

REM		/qb							-	Basic UI Installer
REM		/l*							-	log installation

REM		ISCHECKFORPRODUCTUPDATES	- 	check for updates
REM		SELECT_X					-	Select filetypes for X Product
REM		REGISTER_ALL_MSO_TYPES		-	Register all MS Office types
REM		HideEula					-	Hide the shrink wrap

REM 	ADDLOCAL=ALL 				- 	Add all components
REM		REMOVE xxx,xxx,xx			- 	Individually remove components


msiexec /qb /i LIBREMSIFILE /l* LibO_install_log-%date:~-4,4%%date:~-10,2%%date:~-7,2%-%time:~-11,2%%time:~-8,2%%time:~-5,2%.txt ALLUSERS=1 ISCHECKFORPRODUCTUPDATES=0 SELECT_WORD=1 SELECT_EXCEL=1 SELECT_POWERPOINT=1 REGISTER_ALL_MSO_TYPES=1 HideEula=1 ADDLOCAL=ALL REMOVE=gm_o_Onlineupdate