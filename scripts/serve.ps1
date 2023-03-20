param (
    [string] $Path = "$PSScriptRoot/../site",
    [switch] $BuildUpFront,
    [int] $Port = 4000
)

$jekyllArgs = @()
# if (!$BuildUpFront) {
#     $jekyllArgs += "--skip-initial-build"
# }

$fullPath = (Get-Item $Path).FullName

docker run --rm `
    --publish="4000:$Port" `
    --volume="$($fullPath):/srv/jekyll:Z" `
    --rm --interactive --tty jekyll/jekyll `
    jekyll serve --incremental --watch --force_polling --host 0.0.0.0 --trace @jekyllArgs
