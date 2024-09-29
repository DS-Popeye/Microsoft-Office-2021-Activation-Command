## Office 2021
## Method 1: Using my command line
### Step 1.1: Open cmd program with administrator rights.
- First, you need to open cmd in the admin mode, then run all commands below one by one.
### Step 1.2: Get into the Office directory in cmd.
- For x86 and x64 
```
cd /d %ProgramFiles(x86)%\Microsoft Office\Office16
cd /d %ProgramFiles%\Microsoft Office\Office16
```
- If you install your Office in the ProgramFiles folder, the Office directory depends on the architecture of your OS. If you are not sure of this issue, just run both of the commands above. One of them will be not executed and an error message will be printed on the screen.
### Step 1.3: Install Office 2021 volume license.
```
for /f %x in ('dir /b ..\root\Licenses16\ProPlus2021VL_KMS*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%x"
```
- This step is required. You can not install the KMS client product key of Office without a volume license.
### Step 1.4: Activate your Office using the KMS key.
- Make sure your device is connected to the internet, then run the following commands.
```
cscript ospp.vbs /setprt:1688
cscript ospp.vbs /unpkey:6F7TH >nul
cscript ospp.vbs /inpkey:FXYTK-NJJ8C-GB6DW-3DYQT-6F7TH
cscript ospp.vbs /sethst:e8.us.to
cscript ospp.vbs /act
```
- If you see the error `0xC004F074`, it means that your internet connection is unstable or the server is busy. Please make sure your device is online and try the command “act” again until you succeed.
### Here is all the text you will get in the command prompt window
```
C:\Windows\system32>cd /d %ProgramFiles(x86)%\Microsoft Office\Office16
The system cannot find the path specified.

C:\Windows\system32>cd /d %ProgramFiles%\Microsoft Office\Office16

C:\Program Files\Microsoft Office\Office16>for /f %x in ('dir /b ..\root\Licenses16\ProPlus2021VL_KMS*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%x"

C:\Program Files\Microsoft Office\Office16>cscript ospp.vbs /inslic:"..\root\Licenses16\ProPlus2021VL_KMS_Client_AE-ppd.xrm-ms"
Microsoft (R) Windows Script Host Version 5.812
Copyright (C) Microsoft Corporation. All rights reserved.

---Processing--------------------------
---------------------------------------
Installing Office license: ..\root\licenses16\proplus2021vl_kms_client_ae-ppd.xrm-ms
Office license installed successfully.
---------------------------------------
---Exiting-----------------------------

C:\Program Files\Microsoft Office\Office16>cscript ospp.vbs /inslic:"..\root\Licenses16\ProPlus2021VL_KMS_Client_AE-ul-oob.xrm-ms"
Microsoft (R) Windows Script Host Version 5.812
Copyright (C) Microsoft Corporation. All rights reserved.

---Processing--------------------------
---------------------------------------
Installing Office license: ..\root\licenses16\proplus2021vl_kms_client_ae-ul-oob.xrm-ms
Office license installed successfully.
---------------------------------------
---Exiting-----------------------------

C:\Program Files\Microsoft Office\Office16>cscript ospp.vbs /inslic:"..\root\Licenses16\ProPlus2021VL_KMS_Client_AE-ul.xrm-ms"
Microsoft (R) Windows Script Host Version 5.812
Copyright (C) Microsoft Corporation. All rights reserved.

---Processing--------------------------
---------------------------------------
Installing Office license: ..\root\licenses16\proplus2021vl_kms_client_ae-ul.xrm-ms
Office license installed successfully.
---------------------------------------
---Exiting-----------------------------

C:\Program Files\Microsoft Office\Office16>cscript ospp.vbs /setprt:1688
Microsoft (R) Windows Script Host Version 5.812
Copyright (C) Microsoft Corporation. All rights reserved.

---Processing--------------------------
---------------------------------------
Successfully applied setting.
---------------------------------------
---Exiting-----------------------------

C:\Program Files\Microsoft Office\Office16>cscript ospp.vbs /unpkey:6F7TH >nul

C:\Program Files\Microsoft Office\Office16>cscript ospp.vbs /inpkey:FXYTK-NJJ8C-GB6DW-3DYQT-6F7TH
Microsoft (R) Windows Script Host Version 5.812
Copyright (C) Microsoft Corporation. All rights reserved.

---Processing--------------------------
---------------------------------------
<Product key installation successful>
---------------------------------------
---Exiting-----------------------------

C:\Program Files\Microsoft Office\Office16>cscript ospp.vbs /sethst:e8.us.to
Microsoft (R) Windows Script Host Version 5.812
Copyright (C) Microsoft Corporation. All rights reserved.

---Processing--------------------------
---------------------------------------
Successfully applied setting.
---------------------------------------
---Exiting-----------------------------

C:\Program Files\Microsoft Office\Office16>cscript ospp.vbs /act
Microsoft (R) Windows Script Host Version 5.812
Copyright (C) Microsoft Corporation. All rights reserved.

---Processing--------------------------
---------------------------------------
Installed product key detected - attempting to activate the following product:
SKU ID: cd18ecc0-466a-45d4-8d2e-8c4fa48ae591
LICENSE NAME: Office 21, Office21ProPlus2021R_Grace edition
LICENSE DESCRIPTION: Office 21, RETAIL(Grace) channel
Last 5 characters of installed product key: PG343
ERROR CODE: 0xC004F017
ERROR DESCRIPTION: The Software Licensing Service reported that the license is not installed.
---------------------------------------
Installed product key detected - attempting to activate the following product:
SKU ID: fbdb3e18-a8ef-4fb3-9183-dffd60bd0984
LICENSE NAME: Office 21, Office21ProPlus2021VL_KMS_Client_AE edition
LICENSE DESCRIPTION: Office 21, VOLUME_KMSCLIENT channel
Last 5 characters of installed product key: 6F7TH
<Product activation successful>
---------------------------------------
---------------------------------------
---Exiting-----------------------------

C:\Program Files\Microsoft Office\Office16>
```
## Congratulations! The activation was completed successfully.
#### ==========================================

## Method 2: Using my pre-written batch script
### Step 2.1: Copy the script code below into a new text document.
```
@echo off
title Activate Microsoft Office 2021 (ALL versions) for FREE - office.com&cls&echo =====================================================================================&echo #Project: Activating Microsoft software products for FREE without additional software&echo =====================================================================================&echo.&echo #Supported products:&echo - Microsoft Office Standard 2021&echo - Microsoft Office Professional Plus 2021&echo.&echo.&(if exist "%ProgramFiles%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles%\Microsoft Office\Office16")&(if exist "%ProgramFiles(x86)%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles(x86)%\Microsoft Office\Office16")&(for /f %%x in ('dir /b ..\root\Licenses16\ProPlus2021VL_KMS*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul)&echo.&echo =====================================================================================&echo Activating your product...&cscript //nologo slmgr.vbs /ckms >nul&cscript //nologo ospp.vbs /setprt:1688 >nul&cscript //nologo ospp.vbs /unpkey:6F7TH >nul&set i=1&cscript //nologo ospp.vbs /inpkey:FXYTK-NJJ8C-GB6DW-3DYQT-6F7TH >nul||goto notsupported
:skms
if %i% GTR 10 goto busy
if %i% EQU 1 set KMS=kms7.office.com
if %i% EQU 2 set KMS=e8.us.to
if %i% EQU 3 set KMS=e9.us.to
if %i% GTR 3 goto ato
cscript //nologo ospp.vbs /sethst:%KMS% >nul
:ato
echo =====================================================================================&echo.&echo.&cscript //nologo ospp.vbs /act | find /i "successful" && (echo.&echo =====================================================================================&echo.&echo #Please consider supporting this project: donate to my paypal "devomman@gmail" &echo #Your support is helping me keep my servers running 24/7!&echo.&echo =====================================================================================&choice /n /c YN /m "Would you like to visit my blog [Y,N]?" & if errorlevel 2 exit) || (echo The connection to my KMS server failed! Trying to connect to another one... & echo Please wait... & echo. & echo. & set /a i+=1 & goto skms)
explorer "https://www.office.com/"&goto halt
:notsupported
echo =====================================================================================&echo.&echo Sorry, your version is not supported.&echo.&goto halt
:busy
echo =====================================================================================&echo.&echo Sorry, the server is busy and can't respond to your request. Please try again.&echo.
:halt
pause >nul
```
### Step 2.2: Save this text file as a cmd file. (Eg. office.cmd).
### Step 2.3: Run the cmd file in admin mode.
### Step 2.4: Check the activation status again.
## Done! Your product is activated successfully now.

## More information:

- Here is the KMS client key of Office 2021: ```FXYTK-NJJ8C-GB6DW-3DYQT-6F7TH```.
- The Office 2021 KMS license is valid for 180 days only but can be renewed automatically so you needn’t worry so much about the period.

- If you have any questions or concerns, please comment. I would be glad to explain in more detail. Thank you so much for all your feedback and support!
- If you have any questions please make an issue or discussion.
