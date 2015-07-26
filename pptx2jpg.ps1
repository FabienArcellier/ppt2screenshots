# Script PS1 to convert PPT to screenshot of slides
# Original : https://github.com/utapyngo/pptrasterizer/blob/8b2b9d294af6e0b46c1d764ecd0196e7081e7529/ppt_rasterize.ps1

param(
    [string] $pfilename,
    [string] $output
)

$PSScriptRoot = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition
. $PSScriptRoot\pptx2screenshots.func.ps1

if (-not $pfilename) {
    Write-Host "Usage: powershell -ExecutionPolicy Bypass ""$($script:MyInvocation.MyCommand.Path)"" ""Presentation.pptx"""
    return
}

if (-not (Test-Path $pfilename)) {
    Write-Host "File ""$pfilename"" not found"
    return
}

$application = New-Object -ComObject "PowerPoint.Application"
try {
    Pptx2screenshots $application $pfilename $output "jpg"
}
finally {
    $application.Quit()
}