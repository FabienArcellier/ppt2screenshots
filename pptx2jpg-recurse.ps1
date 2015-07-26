# Script PS1 to convert PPT to screenshot of slides
# Original : https://github.com/utapyngo/pptrasterizer/blob/8b2b9d294af6e0b46c1d764ecd0196e7081e7529/ppt_rasterize.ps1

param(
    [string] $pdirectory,
    [string] $output
)

$PSScriptRoot = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition
. $PSScriptRoot\pptx2screenshots.func.ps1

$application = New-Object -ComObject "PowerPoint.Application"
try {
  $drive_directory= Resolve-Path $pdirectory

  get-childitem -Recurse "$drive_directory\*.pptx" | %{ echo $_.FullName } | % { Pptx2screenshots $application "$_" $output "jpg"}
}
finally {
    $application.Quit()
}