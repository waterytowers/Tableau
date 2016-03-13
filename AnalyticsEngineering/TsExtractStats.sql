SELECT
  title
--, updated_at
--, created_at
--, completed_at
--, started_at
--, site_id
,AVG(EXTRACT(EPOCH FROM (completed_at-created_at))) as "TotalTime_Avg"
,AVG(EXTRACT(EPOCH FROM (completed_at-started_at))) as "RunTime_Avg"
 
,MIN(EXTRACT(EPOCH FROM (completed_at-created_at))) as "TotalTime_Min"
,MIN(EXTRACT(EPOCH FROM (completed_at-started_at))) as "RunTime_Min"
 
,MAX(EXTRACT(EPOCH FROM (completed_at-created_at))) as "TotalTime_Max"
,MAX(EXTRACT(EPOCH FROM (completed_at-started_at))) as "RunTime_Max"
 
 
FROM background_jobs 
WHERE job_name LIKE 'Refresh Extracts'
GROUP BY title

-- date groupings


SELECT
date_trunc('hour',created_at) - INTERVAL '7 hour' as "StartHour"
,COUNT(*) as "Total"
,AVG(EXTRACT(EPOCH FROM (completed_at-created_at))) as "TotalTime_Avg_Sec"
,AVG(EXTRACT(EPOCH FROM (completed_at-started_at))) as "RunTime_Avg_Sec"
FROM background_jobs 
WHERE job_name LIKE 'Refresh Extracts'
GROUP BY
date_trunc('hour',created_at) - INTERVAL '7 hour'
