Set-StrictMode -Version Latest
Set-Location 'C:\TsAdmin\PsOutput'
Import-Module 'C:\Users\<your username>\Documents\WindowsPowerShell\Modules\TsGetExtract\TsGetExtract.psm1'

$connectionString = 'Driver={PostgreSQL ANSI(x64)};Server=<yourserver>; Port=8060; Database=workgroup; Uid=<yourusername>; Pwd=<yourpw>;'
$query = @"
select  distinct
	wb.id as "Workbook_Id",
	wb.name as "Workbook_Name", 
	coalesce(wbv.wb_last_view_time,'1999-01-01') as "wb_last_view_time",
	wb.repository_url as "Workbook_RepositoryURL",
	wb.created_at as "Workbook_CreatedDate",
	wb.updated_at as "Workbook_UpdatedDate",
	wb_su_owner.name as "Workbook_Owner_Username",
	wb_su_owner.friendly_name as "Workbook_Owner_Name",
	p.name as "Workbook_ProjectName",
	wb.size as "Workbook_Size",
	case when coalesce(wb.repository_data_id,1) = 1 then 'TWBX' else 'TWB' end as "Workbook_Type",
	s.name as "Site_Name",
	vs.*,
	current_date as "DateMarker"	
from workbooks wb
left join data_connections dc on dc.owner_id = wb.id and dc.owner_type = 'Workbook'
left join users wb_owner on wb_owner.id = wb.owner_id
left join system_users wb_su_owner on wb_su_owner.id = wb_owner.system_user_id
left join projects p on p.id = wb.project_id
left join sites s on s.id = wb.site_id
left join (
		select distinct
		vs.views_workbook_id as "Workbook_ID",
		w.name as "Workbook_Name",
		w.workbook_url as "Workbook_URL",
		max(vs.last_view_time) over(partition by w.workbook_url) as wb_last_view_time
		from _views_stats vs
		left outer join _workbooks w on w.id = vs.views_workbook_id
	) wbv on wbv."Workbook_URL" = wb.repository_url

left join  ( 
	select views_workbook_id
	, views_url
	, views_name
	, system_users_friendly_name
	, last_view_time
	, nviews from _views_stats
	  ) vs on vs.views_workbook_id = wb.id

"@

Get-TsExtract -connectionString $connectionString -query $query | `
Select-Object -Skip 1 -Property * | `
Export-Csv -Path 'C:\TsAdmin\PsOutput\TsDashboardData.csv' -NoTypeInformation -Delimiter ";" 

##Cleaning for SQL loading##

(Get-Content -Path 'C:\TsAdmin\PsOutput\TsDashboardData.csv') | 
Foreach-Object {$_ -replace "`""," "} | Set-Content 'C:\TsAdmin\PsOutput\TsDashboardData.csv'

