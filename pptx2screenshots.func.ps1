# Convert a presentation into slide screenshot 
#
# $application com object (PowerPoint.Application)
# $pfilename presentation filename
# $output directory to save screenshots
function Pptx2screenshots($application, $pfilename, $output, $format='jpg') {
  try
  {
    $pfilename = Resolve-Path $pfilename
    $filename = Split-Path $pfilename -Leaf
    $name = $filename.substring(0, $filename.lastindexOf("."))
    $slidesPath = Resolve-Path $output
    mkdir $slidesPath -ErrorAction SilentlyContinue | Out-Null
  
    $presentation = $application.Presentations.Open($pfilename, $False, $False, $False)
    try {
      foreach ($original_slide in $presentation.Slides) {
        $i = $original_slide.SlideIndex

        $slide_image_file_name = Join-Path $slidesPath "$name-$i.$format"
        Write-Host "Processing - $slide_image_file_name"
        $original_slide.export($slide_image_file_name, "$format") | Out-Null
      }
    }
    finally {
        $presentation.Close() | Out-Null
    }
  } Catch {
    Write-Error "Fail to manage $pfilename"
  }
}