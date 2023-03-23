
with rpt_parameters_cte1 as 
(
select itemid, cast(parameter as xml) as parameters
from catalog
where type = 2 and path like '%/Blackbaud/AppFx/BBInfinity/System Reports%'
),

rpt_parameters_cte2 as
(
select 
itemid,
y.value('(Name/text())[1]', 'varchar(100)') as param_name,
y.value('(Prompt/text())[1]', 'varchar(100)') as param_prompt,
y.value('(Type/text())[1]', 'varchar(100)') as param_datatype

from rpt_parameters_cte1 x
cross apply x.parameters.nodes('//Parameters/Parameter') Queries (y)
),

rpt_parameters_cte3
as
(
select distinct
x1.itemid,
stuff
((
select distinct cast(' | ' as varchar(max)) + rtrim(ltrim(x2.param_name)) + ' (' + rtrim(ltrim(x2.param_datatype)) + ')'
from rpt_parameters_cte2 x2
where x1.itemid = x2.itemid
for xml path('')
), 1, 1, '') as params

from 
rpt_parameters_cte2 x1
),

reportusers as
(
select 
c.itemid, u.username

from 
policyuserrole pur
join users u on pur.userid = u.userid
join roles r on pur.roleid = r.roleid
join catalog c on pur.policyid = c.policyid

where
c.path like '%power bi reports%'
and r.rolename = 'browser'
),

reportusers_concat as
(
select distinct
x1.itemid,
stuff
((
select distinct cast(',' as varchar(max)) + rtrim(ltrim(x2.username))
from reportusers x2
where x1.itemid = x2.itemid
for xml path('')
), 1, 1, '') as reportusers

from 
reportusers x1
)

select
case when c.path like '%/Blackbaud/AppFx/BBInfinity/System Reports%' then 'SSRS' else 'Power BI' end as reportformat,
c.itemid as reportid,
isnull(c.name, '') name,
isnull(c.path, '') path,
isnull(case when c.path like '%/Blackbaud/AppFx/BBInfinity/System Reports%' then c.description 
else replace(replace(substring(c.description, charindex('[description:',c.description), case when c.description like '%*inactive*%' then charindex('*inactive*',c.description) else 500 end - charindex('[description:',c.description)), '[description: ', ''), ']', '') end, '') as description,
convert(date, c.creationdate) as report_createdate,
convert(date, c.modifieddate) as report_modifieddated,
isnull(p.params, '') params,
isnull(ru.reportusers, '') reportusers,
isnull(case when c.path like '%/Blackbaud/AppFx/BBInfinity/System Reports%' then '' 
else convert(varchar(100), replace(replace(replace(c.description, '[owner: ', ''), ']', ''), ' *inactive*', '')) end, '') reportowner,
isnull(case when c.path like '%/Blackbaud/AppFx/BBInfinity/System Reports%' then ''
when c.description like '%*inactive*%' then 'Inactive' else 'Active' end, '') as reportstatus,
case when c.path like '%/Blackbaud/AppFx/BBInfinity/System Reports%' then '' else convert(varchar(300), 'https://adv114.ad.unc.edu/reports/powerbi' + c.path + '?rs:embed=true') end as reporturl,
isnull(s.name, '') as schedulename,
convert(uniqueidentifier, '095E4CCF-8AC1-432F-ABDE-F66A71FC5C2C') as addedbyid,
convert(uniqueidentifier, '095E4CCF-8AC1-432F-ABDE-F66A71FC5C2C') as changedbyid,
getdate() as dateadded,
getdate() as datechanged

from 
catalog c
left join rpt_parameters_cte3 p on c.itemid = p.itemid 
left join reportusers_concat ru on c.itemid = ru.itemid
left join reportschedule rs on c.itemid = rs.reportid and c.path like '%/Blackbaud/AppFx/BBInfinity/Power BI Reports%' 
left join schedule s on rs.scheduleid = s.scheduleid

where 
c.type in (2, 13)
and 
c.name not like '%old%'
and
(
(c.path like '%/Blackbaud/AppFx/BBInfinity/System Reports/Misc Reports%' and c.name like 'UNC%')
or c.path like '%/Blackbaud/AppFx/BBInfinity/System Reports/UNC Campaign Reports%' 
or c.path like '%/Blackbaud/AppFx/BBInfinity/System Reports/UNC Campaign End%' 
or c.path like '%/Blackbaud/AppFx/BBInfinity/System Reports/UNC Data Checks%' 
or c.path like '%/Blackbaud/AppFx/BBInfinity/System Reports/UNC Financial Reports%' 
or c.path like '%/Blackbaud/AppFx/BBInfinity/System Reports/UNC GAA Reports%' 
or c.path like '%/Blackbaud/AppFx/BBInfinity/System Reports/UNC Grateful Patients%' 
or (c.path like '%/Blackbaud/AppFx/BBInfinity/Power BI Reports%' and c.name like 'unc%' and c.name <> 'UNC_G2' and c.name <> 'UNC_C' and c.name <> 'UNCTicketingDashboard')
)