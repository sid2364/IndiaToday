function Older-Than()
{
PARAM(
[Parameter(ValueFromPipeline=$true)][Object[]]$Files,
[Parameter(ValueFromPipeline=$false, Position=0)][Int]$Days
)
Process
{
$Files | ?{ ( (Get-Date) - $_.LastWriteTime).Days -gt $Days }
}
}