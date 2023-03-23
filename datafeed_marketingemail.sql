


with click as
(
select
constituentid, emailjobrecipientid, emailid,
max(case when sequence = 1 then requestdate end) as click1requestdate,
max(case when sequence = 1 then destination end) as click1destination,
max(case when sequence = 2 then requestdate end) as click2requestdate,
max(case when sequence = 2 then destination end) as click2destination,
max(case when sequence = 3 then requestdate end) as click3requestdate,
max(case when sequence = 3 then destination end) as click3destination,
max(case when sequence = 4 then requestdate end) as click4requestdate,
max(case when sequence = 4 then destination end) as click4destination,
max(case when sequence = 5 then requestdate end) as click5requestdate,
max(case when sequence = 5 then destination end) as click5destination

from
(
select 
map.constituentid,
stats.emailjobrecipientid,
ejr.emailid,
row_number() over (partition by map.constituentid, stats.emailjobrecipientid order by stats.requestdate) as sequence,
stats.requestdate,
case 
when cm.siteid is not null then  '['+ si.shortname + '] ' + isnull(cm.title , '') + ' (' +  stats.url +')'
when sitepages.pagename is not null then sitepages.pagename
else stats.url end as destination
--stats.url

from 
dbo.stats
join emailjob_recipient ejr on stats.emailjobrecipientid=ejr.id 
left join dbo.sitepages on stats.pageid=sitepages.id 
join usr_unc_marketingeffortrecipientmap map on ejr.id=map.emailjobrecipientid
left join usr_unc_unc_communicationurlmapping cm on cast(substring(stats.[url] , charindex('=', cm.url)+1, len(cm.url)) as varchar(100)) = cm.url
left join [site] si on si.id = cm.siteid

where
stats.sourceid > 0
and isnull(sitepages.pagename,'') not like '%unsub%' 
and isnull(sitepages.pagename,'') not like '%privacy%'
) x

group by
constituentid, emailjobrecipientid, emailid
)

select distinct
c.id as constituentid, 
ejr.id as emailjobrecipientid, 
ejr.emailaddress as recipientemail,
e.name as emailname, 
ejr.sent, 
ejr.dsned as bounced, 
ejr.opened, 
ejr.sentdate, 
ejr.openeddate, 
map.marketingeffortid,
s.id as siteid,
s.shortname as siteshortname,
s.name as sitename,
click.click1requestdate, click.click1destination,
click.click2requestdate, click.click2destination,
click.click3requestdate, click.click3destination,
click.click4requestdate, click.click4destination,
click.click5requestdate, click.click5destination,
unsubscribe.unsubscribed,
unsubscribe.unsubscribe_date

from
constituent c
join usr_unc_marketingeffortrecipientmap map on c.id = map.constituentid 
join emailjob_recipient ejr on map.emailjobrecipientid = ejr.id 
join email e on ejr.emailid = e.id
join mktsegmentation m on map.marketingeffortid = m.id
join site s on m.siteid = s.id
left join click on c.id = click.constituentid and ejr.id = click.emailjobrecipientid

outer apply 
(
select top 1 1 as unsubscribed, a.dateadded as unsubscribe_date
from usr_unc_bbisanonymousunsubscribe a
where a.trid = ejr.mergeid
) unsubscribe

where
ejr.sent = 1

