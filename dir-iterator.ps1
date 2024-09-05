param (
    [string]$DirectoryPath,
    [string]$Program
)

ForEach ($File in (Get-ChildItem -Path $DirectoryPath)) {
    python $Program "$DirectoryPath/$File"
}