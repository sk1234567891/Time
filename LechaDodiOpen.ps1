Stop-Process -Name POWERPNT -ErrorAction SilentlyContinue
Start-Process powerpnt -ArgumentList ("/s " + "$PSScriptRoot\pp\LechaDodi.pptx")