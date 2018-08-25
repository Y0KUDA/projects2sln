Param([parameter(mandatory)][string]$path,
	[string]$sln="Solution",
    [bool]$visible=$False,
    [string]$filter = "*.csproj,*.vbproj,*.vcxproj")
$defaultLoc=Get-Location
$f=$filter -split ","
$path=Resolve-Path($path)
function _CreateSlnRecurse($CurDir){
    Get-ChildItem $f -File |%{
        Start-Sleep -Milliseconds 100
        $CurDir.AddFromFile($_.FullName)
    }
	Get-ChildItem -Directory |%{
		Set-Location $_.Name
        Start-Sleep -Milliseconds 100
        $CurDirTmp= $CurDir.AddSolutionFolder($_.Name)
        Start-Sleep -Milliseconds 100
		_CreateSlnRecurse($CurDirTmp.Object)
		Set-Location ..
	}
}

#$VSDTE = [Runtime.InteropServices.Marshal]::GetActiveObject("VisualStudio.DTE")
$VSDTE=New-Object -ComObject VisualStudio.DTE
$VSDTE.MainWindow.Visible=$visible
$Sol=$VSDTE.Solution
$VSDTE.Solution.Create($defaultLoc,$sln)
Set-Location $path
_CreateSlnRecurse($Sol)
$VSDTE.Solution.SaveAs($sln)
$VSDTE.Quit()
Set-Location $defaultLoc






