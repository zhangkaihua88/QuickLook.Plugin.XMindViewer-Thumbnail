Remove-Item ..\QuickLook.Plugin.XMindViewer.qlplugin -ErrorAction SilentlyContinue

$files = Get-ChildItem -Path ..\bin\Release\ -Exclude *.pdb,*.xml
Compress-Archive $files ..\QuickLook.Plugin.XMindViewer.zip
Move-Item ..\QuickLook.Plugin.XMindViewer.zip ..\QuickLook.Plugin.XMindViewer.qlplugin