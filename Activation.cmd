@cls
@echo off

fsutil>nul
if errorlevel 1 (
  echo 请以管理员权限运行本脚本！
  echo PS：右键击本脚本，并点击“以管理员身份运行（A）”。
  pause>nul
  exit
)

::如果你有自建了KMS激活服务器，记得把localhost替换为你的服务器地址域名哦！！
set KMS=localhost

ping -n 1 %KMS%>nul
if errorlevel 1 (
  echo KMS服务器暂时离线，无法激活！
  echo 请确保你已正确联网，如果还是出现本错误，请改天再来。
  pause>nul
  exit
)

echo 正在激活Windows，请稍后 ...
cscript slmgr.vbs -skms %KMS%>nul
cscript slmgr.vbs -ato>nul

echo 正在查找安装的Office并激活之，请稍后 ...
set success=false

setlocal enabledelayedexpansion
FOR /D %%I IN ("%ProgramFiles%\Microsoft Office\Office*" "%ProgramFiles(x86)%\Microsoft Office\Office*") DO (
  set OSPP="%%I\OSPP.VBS"
  IF EXIST !OSPP! (
    echo 找到了Offce安装路径%%I，正在激活 ...
    cscript !OSPP! /sethst:%KMS%>nul
    cscript !OSPP! /act>nul
    set success=true
  )
)

if %success%==false (
  echo 并没有找到任何Microsoft Office安装！
  echo 不需要激活Office请直接按回车继续脚本。
  echo 请键入安装路径，例如“C:\Program Files\Microsoft Office\Office16”：

  set /p OSPP=

  set OSPP="!OSPP!\OSPP.VBS"
  IF EXIST !OSPP! (
    echo 找到Office安装！
    cscript !OSPP! /sethst:%KMS%>nul
    cscript !OSPP! /act>nul
  )
)

echo 工作完成！
echo 如果你发现Windows还没激活，请先安装适合的密钥，然后再使用本脚本。
echo 安装密钥的命令：cscript slmgr.vbs /ilc XXXXX-XXXXX-XXXXX-XXXXX-XXXXX
echo 密钥可以在这里找到：https://technet.microsoft.com/en-us/library/jj612867.aspx
echo 有任何错误或者疑问，请联系我weiqi_chen@outlook.com

pause>nul