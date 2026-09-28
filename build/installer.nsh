; NSIS uses the final component of InstallDir for its native Browse dialog.
; Setting $INSTDIR at runtime alone does not configure that browse suffix.
!macro customHeader
  InstallDir "$LOCALAPPDATA\Programs\local-docs"
!macroend

; Delete only the known application payload. Never recursively remove $INSTDIR:
; the default library lives in $INSTDIR\data and must survive upgrade/uninstall.
!macro customRemoveFiles
  SetOutPath "$TEMP"
  ClearErrors
  Delete "$INSTDIR\${APP_EXECUTABLE_FILENAME}"
  IfErrors 0 +2
    Abort "Please close Local Docs before continuing. Your documents have not been removed."
  Delete "$INSTDIR\chrome_100_percent.pak"
  Delete "$INSTDIR\chrome_200_percent.pak"
  Delete "$INSTDIR\d3dcompiler_47.dll"
  Delete "$INSTDIR\dxcompiler.dll"
  Delete "$INSTDIR\dxil.dll"
  Delete "$INSTDIR\ffmpeg.dll"
  Delete "$INSTDIR\icudtl.dat"
  Delete "$INSTDIR\LICENSE.electron.txt"
  Delete "$INSTDIR\LICENSES.chromium.html"
  Delete "$INSTDIR\resources.pak"
  Delete "$INSTDIR\snapshot_blob.bin"
  Delete "$INSTDIR\v8_context_snapshot.bin"
  Delete "$INSTDIR\vk_swiftshader_icd.json"
  Delete "$INSTDIR\vk_swiftshader.dll"
  Delete "$INSTDIR\vulkan-1.dll"
  Delete "$INSTDIR\uninstallerIcon.ico"
  Delete "$INSTDIR\${UNINSTALL_FILENAME}"
  RMDir /r "$INSTDIR\locales"
  RMDir /r "$INSTDIR\resources"
  ; Only removes the installation directory if it is empty. Data is retained.
  RMDir "$INSTDIR"
  ClearErrors
!macroend
