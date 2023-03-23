

---------------------------------------------------------------
-- matchdata (schedule) info
---------------------------------------------------------------
with matchdata_cte1 as 
(
select subscriptionid, cast(matchdata as xml) as matchdata,
case 
when matchdata like '%DailyRecurrence%' then 'Daily'
when matchdata like '%WeeklyRecurrence%' then 'Weekly'
when matchdata like '%MonthlyRecurrence%' then 'Monthly'
when matchdata like '%MonthlyDOWRecurrence%' then 'Monthly Day of Week'
when matchdata like '%MinuteRecurrence%' then 'Hourly'
else 'On Demand' end as recurrence

from subscriptions
),

matchdata_cte2 as
(
select 
x.subscriptionid, 
x.recurrence,
x.matchdata.value('(*:ScheduleDefinition/*:StartDateTime/node())[1]', 'varchar(30)') runtime,
x.matchdata.value('(*:ScheduleDefinition/*:EndDate/node())[1]', 'varchar(30)') enddate,
x.matchdata.value('(*:ScheduleDefinition/*:DailyRecurrence/*:DaysInterval/node())[1]', 'varchar(30)') dayinterval,
x.matchdata.value('(*:ScheduleDefinition/*:WeeklyRecurrence/*:WeeksInterval/node())[1]', 'varchar(30)') weekinterval,
x.matchdata.value('(*:ScheduleDefinition/*:WeeklyRecurrence/*:DaysOfWeek/*:Sunday/node())[1]', 'varchar(30)') weekday_sun,
x.matchdata.value('(*:ScheduleDefinition/*:WeeklyRecurrence/*:DaysOfWeek/*:Monday/node())[1]', 'varchar(30)') weekday_mon,
x.matchdata.value('(*:ScheduleDefinition/*:WeeklyRecurrence/*:DaysOfWeek/*:Tuesday/node())[1]', 'varchar(30)') weekday_tue,
x.matchdata.value('(*:ScheduleDefinition/*:WeeklyRecurrence/*:DaysOfWeek/*:Wednesday/node())[1]', 'varchar(30)') weekday_wed,
x.matchdata.value('(*:ScheduleDefinition/*:WeeklyRecurrence/*:DaysOfWeek/*:Thursday/node())[1]', 'varchar(30)') weekday_thu,
x.matchdata.value('(*:ScheduleDefinition/*:WeeklyRecurrence/*:DaysOfWeek/*:Friday/node())[1]', 'varchar(30)') weekday_fri,
x.matchdata.value('(*:ScheduleDefinition/*:WeeklyRecurrence/*:DaysOfWeek/*:Saturday/node())[1]', 'varchar(30)') weekday_sat,
x.matchdata.value('(*:ScheduleDefinition/*:MonthlyRecurrence/*:Days/node())[1]', 'varchar(30)') monthdays,
x.matchdata.value('(*:ScheduleDefinition/*:MonthlyRecurrence/*:MonthsOfYear/*:January/node())[1]', 'varchar(5)')  month_jan,
x.matchdata.value('(*:ScheduleDefinition/*:MonthlyRecurrence/*:MonthsOfYear/*:February/node())[1]', 'varchar(5)') month_feb,
x.matchdata.value('(*:ScheduleDefinition/*:MonthlyRecurrence/*:MonthsOfYear/*:March/node())[1]', 'varchar(5)') month_mar,
x.matchdata.value('(*:ScheduleDefinition/*:MonthlyRecurrence/*:MonthsOfYear/*:April/node())[1]', 'varchar(5)') month_apr,
x.matchdata.value('(*:ScheduleDefinition/*:MonthlyRecurrence/*:MonthsOfYear/*:May/node())[1]', 'varchar(5)') month_may,
x.matchdata.value('(*:ScheduleDefinition/*:MonthlyRecurrence/*:MonthsOfYear/*:June/node())[1]', 'varchar(5)') month_jun,
x.matchdata.value('(*:ScheduleDefinition/*:MonthlyRecurrence/*:MonthsOfYear/*:July/node())[1]', 'varchar(5)') month_jul,
x.matchdata.value('(*:ScheduleDefinition/*:MonthlyRecurrence/*:MonthsOfYear/*:August/node())[1]', 'varchar(5)') month_aug,
x.matchdata.value('(*:ScheduleDefinition/*:MonthlyRecurrence/*:MonthsOfYear/*:September/node())[1]', 'varchar(5)') month_sep,
x.matchdata.value('(*:ScheduleDefinition/*:MonthlyRecurrence/*:MonthsOfYear/*:October/node())[1]', 'varchar(5)') month_oct,
x.matchdata.value('(*:ScheduleDefinition/*:MonthlyRecurrence/*:MonthsOfYear/*:November/node())[1]', 'varchar(5)') month_nov,
x.matchdata.value('(*:ScheduleDefinition/*:MonthlyRecurrence/*:MonthsOfYear/*:December/node())[1]', 'varchar(5)') month_dec,
x.matchdata.value('(*:ScheduleDefinition/*:MonthlyDOWRecurrence/*:WhichWeek/node())[1]', 'varchar(10)') month_dow_week,
x.matchdata.value('(*:ScheduleDefinition/*:MonthlyDOWRecurrence/*:DaysOfWeek/*:Sunday/node())[1]', 'varchar(5)') month_dow_sun,
x.matchdata.value('(*:ScheduleDefinition/*:MonthlyDOWRecurrence/*:DaysOfWeek/*:Monday/node())[1]', 'varchar(5)') month_dow_mon,
x.matchdata.value('(*:ScheduleDefinition/*:MonthlyDOWRecurrence/*:DaysOfWeek/*:Tuesday/node())[1]', 'varchar(5)') month_dow_tue,
x.matchdata.value('(*:ScheduleDefinition/*:MonthlyDOWRecurrence/*:DaysOfWeek/*:Wednesday/node())[1]', 'varchar(5)') month_dow_wed,
x.matchdata.value('(*:ScheduleDefinition/*:MonthlyDOWRecurrence/*:DaysOfWeek/*:Thursday/node())[1]', 'varchar(5)') month_dow_thu,
x.matchdata.value('(*:ScheduleDefinition/*:MonthlyDOWRecurrence/*:DaysOfWeek/*:Friday/node())[1]', 'varchar(5)') month_dow_fri,
x.matchdata.value('(*:ScheduleDefinition/*:MonthlyDOWRecurrence/*:DaysOfWeek/*:Saturday/node())[1]', 'varchar(5)') month_dow_sat

from 
matchdata_cte1 x
cross apply x.matchdata.nodes('//*:ScheduleDefinition') Queries (y)
),


---------------------------------------------------------------
-- extensionsettings (delivery) info
---------------------------------------------------------------
extensionsettings_cte1 as 
(
select subscriptionid, cast(extensionsettings as xml) as extensionsettings
from subscriptions
),

extensionsettings_cte2 as
(
select 
subscriptionid,
isnull(y.value('(./*:Name/text())[1]', 'nvarchar(1024)'),'Value') as settingname,
y.value('(./*:Value/text())[1]', 'nvarchar(max)') as settingvalue

from 
extensionsettings_cte1 x
cross apply x.extensionsettings.nodes('//*:ParameterValue') Queries (y)
),

extensionsettings_cte3 as
(
select 
subscriptionid,
max(case when settingname = 'to' then settingvalue else null end) as ext_to,
max(case when settingname = 'cc' then settingvalue else null end) as ext_cc,
max(case when settingname = 'replyto' then settingvalue else null end) as ext_replyto,
max(case when settingname = 'IncludeReport' then settingvalue else null end) as ext_includereport,
max(case when settingname = 'RenderFormat' then settingvalue else null end) as ext_renderformat,
max(case when settingname = 'Priority' then settingvalue else null end) as ext_priority,
max(case when settingname = 'Subject' then settingvalue else null end) as ext_subject,
max(case when settingname = 'IncludeLink' then settingvalue else null end) as ext_includelink,
max(case when settingname = 'Path' then settingvalue else null end) as ext_path,
max(case when settingname = 'Filename' then settingvalue else null end) as ext_filename,
max(case when settingname = 'RENDER_FORMAT' then settingvalue else null end) as ext_fileshare_renderformat

from 
extensionsettings_cte2

group by
subscriptionid
),

---------------------------------------------------------------
-- parameter info
---------------------------------------------------------------
parameters_cte1 as 
(
select subscriptionid, cast(parameters as xml) as parameters
from subscriptions
),

parameters_cte2 as
(
select 
subscriptionid,
isnull(y.value('(./*:Name/text())[1]', 'nvarchar(1024)'),'Value') + ': ' +
y.value('(./*:Value/text())[1]', 'nvarchar(max)') as paramater

from parameters_cte1 x
cross apply x.parameters.nodes('//*:ParameterValue') Queries (y)
),

parameters_cte3
as
(
select 
subscriptionid,
paramater,
row_number() over(partition by subscriptionid order by paramater) as seq

from
parameters_cte2

where 
paramater is not null
),

---------------------------------------------------------------
-- dataset info
---------------------------------------------------------------
datasettings_cte1 as 
(
select subscriptionid, cast(datasettings as xml) as datasettings
from subscriptions
where datasettings is not null 
),

datasettings_cte2 as
(
select 
subscriptionid, 
x.datasettings.value('(*:DataSet/*:Fields/*:Field/*:Alias/node())[1]', 'varchar(30)') alias,
x.datasettings.value('(*:DataSet/*:Query/*:CommandText/node())[1]', 'varchar(max)') query

from 
datasettings_cte1 x
cross apply x.datasettings.nodes('//*:DataSet') Queries (y)
)

select
c.itemid as reportid,
c.name,
c.path,
c.description,
c.creationdate as report_createdate,
u1.username as report_createdby,
c.modifieddate report_modifieddate,
u2.username as report_modifiedby,
u3.username as subscription_createdby,
s.modifieddate subscription_modifieddate,
u4.username as subscription_modifiedby,
s.description as subscription_description,
s.subscriptionid,
s.laststatus,
s.eventtype,
s.lastruntime,
replace(s.deliveryextension, 'Report Server ', '') as deliveryextension,
case when s.inactiveflags = 0 then 'Enabled' else 'Disabled' end as enabled,

case
when s.inactiveflags = 0 and isnull(md.enddate, '2099-01-01') > getdate() then 'Active'
when s.inactiveflags <> 0 then 'Disabled'
when isnull(md.enddate, '2099-01-01') < getdate() then 'Expired'
end as subscription_status,

case
when s.laststatus like '%not valid%' then 1
when s.laststatus like '%1 error%' then 1
when s.laststatus like '%2 error%' then 1
when s.laststatus like '%3 error%' then 1
when s.laststatus like '%4 error%' then 1
when s.laststatus like '%5 error%' then 1
when s.laststatus like '%6 error%' then 1
when s.laststatus like '%7 error%' then 1
when s.laststatus like '%8 error%' then 1
when s.laststatus like '%9 error%' then 1
when s.laststatus like '%fail%' then 1
when s.laststatus like 'error:%' then 1
else 0 end as last_status_error,

md.recurrence,
convert(varchar(15), cast(md.runtime as time),100) runtime,
md.enddate,

stuff
(
  coalesce(','+ nullif(case when md.weekday_sun = 'true' then 'Sun' else null end, ''), '')
+ coalesce(','+ nullif(case when md.weekday_mon = 'true' then 'Mon' else null end, ''), '')
+ coalesce(','+ nullif(case when md.weekday_tue = 'true' then 'Tue' else null end, ''), '')
+ coalesce(','+ nullif(case when md.weekday_wed = 'true' then 'Wed' else null end, ''), '')
+ coalesce(','+ nullif(case when md.weekday_thu = 'true' then 'Thu' else null end, ''), '')
+ coalesce(','+ nullif(case when md.weekday_fri = 'true' then 'Fri' else null end, ''), '')
+ coalesce(','+ nullif(case when md.weekday_sat = 'true' then 'Sat' else null end, ''), '')
,1,1,'') as weekday_list,

stuff
(
  coalesce(','+ nullif(case when md.month_jan = 'true' then 'Jan' else null end, ''), '')
+ coalesce(','+ nullif(case when md.month_feb = 'true' then 'Feb' else null end, ''), '')
+ coalesce(','+ nullif(case when md.month_mar = 'true' then 'Mar' else null end, ''), '')
+ coalesce(','+ nullif(case when md.month_apr = 'true' then 'Apr' else null end, ''), '')
+ coalesce(','+ nullif(case when md.month_may = 'true' then 'May' else null end, ''), '')
+ coalesce(','+ nullif(case when md.month_jun = 'true' then 'Jun' else null end, ''), '')
+ coalesce(','+ nullif(case when md.month_jul = 'true' then 'Jul' else null end, ''), '')
+ coalesce(','+ nullif(case when md.month_aug = 'true' then 'Aug' else null end, ''), '')
+ coalesce(','+ nullif(case when md.month_sep = 'true' then 'Sep' else null end, ''), '')
+ coalesce(','+ nullif(case when md.month_oct = 'true' then 'Oct' else null end, ''), '')
+ coalesce(','+ nullif(case when md.month_nov = 'true' then 'Nov' else null end, ''), '')
+ coalesce(','+ nullif(case when md.month_dec = 'true' then 'Dec' else null end, ''), '')
,1,1,'') as month_list,

replace(md.month_dow_week, 'Week', '')  + ' Week: ' + 
stuff
(
+ coalesce(','+ nullif(case when md.month_dow_sun = 'true' then 'Sun' else null end, ''), '')
+ coalesce(','+ nullif(case when md.month_dow_mon = 'true' then 'Mon' else null end, ''), '')
+ coalesce(','+ nullif(case when md.month_dow_tue = 'true' then 'Tue' else null end, ''), '')
+ coalesce(','+ nullif(case when md.month_dow_wed = 'true' then 'Wed' else null end, ''), '')
+ coalesce(','+ nullif(case when md.month_dow_thu = 'true' then 'Thu' else null end, ''), '')
+ coalesce(','+ nullif(case when md.month_dow_fri = 'true' then 'Fri' else null end, ''), '')
+ coalesce(','+ nullif(case when md.month_dow_sat = 'true' then 'Sat' else null end, ''), '')
,1,1,'') as month_dow_list,

x.ext_to,
x.ext_cc,
x.ext_replyto,
x.ext_includereport,
case when x.ext_renderformat = 'EXCELOPENXML' then 'EXCEL' else x.ext_renderformat end as ext_renderformat,
case when x.ext_fileshare_renderformat = 'EXCELOPENXML' then 'EXCEL' else x.ext_fileshare_renderformat end as ext_fileshare_renderformat,
x.ext_priority,
replace(x.ext_subject, '@ReportName', c.name) as ext_subject,
x.ext_includelink,
x.ext_path,
x.ext_filename,
stuff
(
  coalesce(','+ nullif(p1.paramater, ''), '')
+ coalesce(','+ nullif(p2.paramater, ''), '')
+ coalesce(','+ nullif(p3.paramater, ''), '')
+ coalesce(','+ nullif(p4.paramater, ''), '')
+ coalesce(','+ nullif(p5.paramater, ''), '')
+ coalesce(','+ nullif(p6.paramater, ''), '')
,1,1,'') as paramaterlist,

d.alias,
d.query,
case 
when charindex('USR_UNC_USR_UNC_GETLISTRECIPS', d.query) > 0
then substring(d.query, charindex('USR_UNC_USR_UNC_GETLISTRECIPS', d.query) + 31, 36)
end as mailinglist

from 
catalog c
join subscriptions s on c.itemid = s.report_oid
left join users u1 on c.createdbyid = u1.userid
left join users u2 on c.modifiedbyid = u2.userid
left join users u3 on s.ownerid = u3.userid
left join users u4 on c.modifiedbyid = u4.userid
join matchdata_cte2 md on s.subscriptionid = md.subscriptionid
left join extensionsettings_cte3 x on s.subscriptionid = x.subscriptionid
left join parameters_cte3 p1 on s.subscriptionid = p1.subscriptionid and p1.seq = 1
left join parameters_cte3 p2 on s.subscriptionid = p2.subscriptionid and p2.seq = 2 
left join parameters_cte3 p3 on s.subscriptionid = p3.subscriptionid and p3.seq = 3
left join parameters_cte3 p4 on s.subscriptionid = p4.subscriptionid and p4.seq = 4
left join parameters_cte3 p5 on s.subscriptionid = p5.subscriptionid and p5.seq = 5
left join parameters_cte3 p6 on s.subscriptionid = p6.subscriptionid and p6.seq = 6
left join datasettings_cte2 d on s.subscriptionid = d.subscriptionid

where
c.path like '/Blackbaud/AppFx/BBInfinity/System Reports/%'
and s.eventtype = 'TimedSubscription'
and s.subscriptionid <> '4250213F-3AF3-43AC-846C-99C488B18013'

