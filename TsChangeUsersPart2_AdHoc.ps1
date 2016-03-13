#Requires -Version 3
Set-StrictMode -Version Latest
$Kenobi = (get-content "<file path to your config file which is just a key/value list>") -join "`n" | convertfrom-stringdata


$workpath = <set your workpath here>


#### nothing to change from this point on (unless you want to change your sql query) ###

if (Test-Path $workpath) {
    Set-Location $workpath
    } else {
        mkdir $workpath ; Set-Location $workpath
       }

$ExtractDate = (Get-Date).ToString('yyyy-MM-dd')

# step 1: get users via postgres

$UserQry = @"
SELECT DISTINCT
 u.name as "username"
 ,s.luid as "site-id"
 ,case s.name
   <put the site logic here b/c the rest api will need site name>
   else s.name end  as "site-name"
 ,uu.luid as "user-id"
 ,elUser.pw
FROM users uu
INNER JOIN _users u on u.system_user_id = uu.system_user_id
  INNER JOIN sites s on s.id = uu.site_id
    JOIN (
		select distinct
		z.username as "username"
		,md5(random()::text) as "pw"
		from(
			select distinct
			pw.name as "username"
			from _users pw
		) z
	) elUser on elUser.username = u.name
-- WHERE <limit for some business reason>
"@
$UserConn = "Driver=$($Kenobi.StarKillerDriver);Server=$($Kenobi.StarkillerServer); Port=$($Kenobi.StarKillerPort); Database=$($Kenobi.StarkillerDatabase); Uid=$($Kenobi.StarkillerUid); Pwd=$($Kenobi.StarkillerPwd);"
$Users=Get-TsExtract -connectionString $UserConn -query $UserQry | select -skip 1 -property username,site-id,site-name,user-id,pw

# Step 2: Update users

foreach($user in $Users) {
$authtoken=$null
$SignIn =[xml]@"
<tsRequest> 
 <credentials name="$($Kenobi.VaderUsername)" password="$($Kenobi.VaderPassword)" >
 <site contentUrl="$($user.'site-name')" />
 </credentials> 
</tsRequest>
"@

$tsResponse = irm -uri "$($Kenobi.VaderRESTSignin)" -Method Post -Body $SignIn
$authtoken = $tsResponse.tsResponse.credentials.token

[xml]$tsPayload=@"
<tsRequest>
 <user     
    password="$($user.pw)"
    />
</tsRequest>
"@
    irm -Uri "$($Kenobi.VaderRESTSignin.Replace('auth/signin',''))sites/$($user.'site-id')/users/$($user.'user-id')" -Headers @{"X-Tableau-Auth"=$authtoken} -Method PUT -Body $tsPayload -OutVariable UserAdds
    $UserAdds.tsresponse.user | select name,fullname,siteRole,@{n='pw';e={$user.'pw'}},@{n='siteName';e={$user.'site-name'}}, `
    @{n='ExtractDate';e={$extractdate}} | export-csv "$($workpath)\EmpsAdds.csv" -Delimiter ";" -NoTypeInformation -Append
}

#step 3 - sign-out

irm -Uri "$($Kenobi.VaderRESTSignin.Replace('signin','signout'))" -Method POST -Headers @{"X-Tableau-Auth"=$authtoken}