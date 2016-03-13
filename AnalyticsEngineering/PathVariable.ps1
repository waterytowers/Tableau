$version = "<Tableau Server Version>"
$scope = "machine"
$oldPath = @([Environment]::GetEnvironmentVariable(“Path”,$scope) –split ";")
$oldPath += "C:\Program Files\Tableau\Tableau Server\$($version)\bin"
$newPath = $oldPath –join ";"
[Environment]::SetEnvironmentVariable(“Path”,$newPath,$scope)
