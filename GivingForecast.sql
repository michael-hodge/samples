



drop table if exists #commitment_tx;  --
drop table if exists #imported;
drop table if exists #importedbymonth;  --
drop table if exists #melt;
drop table if exists #melt_pct;
drop table if exists #pipeline;
drop table if exists #forecastdata;
drop table if exists #slope;
drop table if exists #avg;
drop table if exists #intercept;
drop table if exists #agforecast;

declare @majorgiving money = 100000;
declare @principalgiving money = 5000000;


-------------------------------------------------------------------------------
-- commitment transactions
-------------------------------------------------------------------------------

-- gifts, pledges, grants
select *
into #commitment_tx
from 
(
select
'Commitments: Transactions' source,
cal.fiscalyear,
cal.firstofmonth date,
cal.fy_month_num fymonth,
s.shortname as site,
s.name as sitename,
case 
when s.shortname in ('LCCC','MED', 'MEDF') then 'UNC Health'
when s.shortname in ('ADMU','LTNX', 'CHPDP', 'CHAN', 'CRC', 'FPG', 'HSRC', 'IPRC', 'OVCR', 'PROV', 'RENCI', 'SHSC') then 'Units Without Development Offices'
else s.name end as sitegroup,
--d.designationnumber,
--d.description as designation,
case 
when r.transactionamount >= @principalgiving then 'Principal Giving'
when r.transactionamount >= @majorgiving then 'Major Giving'
else 'Annual Giving' end as gift_level,
case
when r.transactiontype = 'payment' and r.application = 'matching gift' then 'Matching Gift Payment'
when r.transactiontype = 'payment' and r.application = 'other'  then 'Grant'
when r.transactiontype = 'payment' and r.application not in  ('matching gift', 'other') then 'Gift'
when r.transactiontype = 'pledge' then 'Pledge' 
end as gift_type,
--sum(r.transactionamount) amount,
--sum(r.transactionamount) amount,
sum(r.transactionamount * rand()*(2-.25)) amount

from
revenue r
join constituent c on r.constituentid = c.id
join sites s on r.siteid = s.id
join designations d on r.designationid = d.id
join calendar cal on r.transactiondate = cal.date

where
s.siteid not in (' ', '35', '77', '90', '99')
and 
(
-- gift
(r.transactiontype = 'payment' and r.applicationtype = 'gift' and r.application not in ('pledge', 'other', 'planned gift', 'membership')) or 
-- pledge
(r.transactiontype = 'pledge') or  
-- Grant
(r.grantpayment = 'yes') 
)
and r.transactiondate between '2013-07-01' and getdate()

group by
cal.fiscalyear,
cal.firstofmonth,
cal.fy_month_num,
s.shortname,
s.name,
case 
when s.shortname in ('LCCC','MED', 'MEDF') then 'UNC Health'
when s.shortname in ('ADMU','LTNX', 'CHPDP', 'CHAN', 'CRC', 'FPG', 'HSRC', 'IPRC', 'OVCR', 'PROV', 'RENCI', 'SHSC') then 'Units Without Development Offices'
else ' ' + s.name end,
--d.designationnumber,
--d.description,
case 
when r.transactionamount >= @principalgiving then 'Principal Giving'
when r.transactionamount >= @majorgiving then 'Major Giving'
else 'Annual Giving' end,
case
when r.transactiontype = 'payment' and r.application = 'matching gift' then 'Matching Gift Payment'
when r.transactiontype = 'payment' and r.application = 'other'  then 'Grant'
when r.transactiontype = 'payment' and r.application not in  ('matching gift', 'other') then 'Gift'
when r.transactiontype = 'pledge' then 'Pledge' 
end

union

-- planned gifts
select
'Commitments: Transactions' source,
cal.fiscalyear,
cal.firstofmonth date,
cal.fy_month_num fymonth,
s.shortname as site,
s.name as sitename,
case 
when s.shortname in ('LCCC','MED', 'MEDF') then 'UNC Health'
when s.shortname in ('ADMU','LTNX', 'CHPDP', 'CHAN', 'CRC', 'FPG', 'HSRC', 'IPRC', 'OVCR', 'PROV', 'RENCI', 'SHSC') then 'Units Without Development Offices'
else s.name end as sitegroup,
--d.designationnumber,
--d.description as designation,
case 
when pg.giftamount >= @principalgiving then 'Principal Giving'
when pg.giftamount >= @majorgiving then 'Major Giving'
else 'Annual Giving' end as gift_level,
'Planned Gift' as gift_type,
--sum(pg.giftamount) amount,
--sum(pg.giftamount) amount,
sum(pg.giftamount * rand()*(2-.25)) amount

from 
plannedgifts pg
join sites s on pg.site = s.shortname
join designations d on pg.designationid = d.id
join calendar cal on pg.giftdate = cal.date

where 
pg.isrevocable in (0, 1)
and pg.status = 'accepted'
and pg.iscontingent = 0
and s.siteid not in (' ', '35', '77', '90', '99')
and pg.giftdate between '2013-07-01' and getdate()

group by
cal.fiscalyear,
cal.firstofmonth,
cal.fy_month_num,
s.shortname,
s.name,
case 
when s.shortname in ('LCCC','MED', 'MEDF') then 'UNC Health'
when s.shortname in ('ADMU','LTNX', 'CHPDP', 'CHAN', 'CRC', 'FPG', 'HSRC', 'IPRC', 'OVCR', 'PROV', 'RENCI', 'SHSC') then 'Units Without Development Offices'
else ' ' + s.name end,
--d.designationnumber,
--d.description,
case 
when pg.giftamount >= @principalgiving then 'Principal Giving'
when pg.giftamount >= @majorgiving then 'Major Giving'
else 'Annual Giving' end
)x;


-------------------------------------------------------------------------------
-- annual giving forecast
-------------------------------------------------------------------------------

select fiscalyear, fymonth, site, sitename, sitegroup, sum(amount) amount
into #forecastdata
from #commitment_tx
where fiscalyear between 2013 and 2024 and gift_level = 'annual giving' 
group by fiscalyear, fymonth, site, sitename, sitegroup
order by fiscalyear, fymonth, site, sitename, sitegroup


-- get slope
select
site, sitename, sitegroup, fymonth,
case when (n * sum_x2 - sum_x * sum_x) = 0 then 0 else (n * sum_xy - sum_x * sum_y)/(n * sum_x2 - sum_x * sum_x) end as slope
into #slope
from
(
select
site, sitename, sitegroup, fymonth,
count(*) as n,
sum(fiscalyear) as sum_x,
sum(fiscalyear * fiscalyear) as sum_x2,
sum(amount) as sum_y,
sum(fiscalyear * amount) as sum_xy
from #forecastdata h 
group by site, sitename, sitegroup, fymonth
)x
order by site, sitename, sitegroup, fymonth


-- get intercept
select site, sitename, sitegroup, fymonth, avg(fiscalyear) as avg_x, avg(amount) as avg_y
into #avg
from #forecastdata 
group by site, sitename, sitegroup, fymonth;
 
select a.site, a.sitename, a.sitegroup, a.fymonth, (a.avg_y - (s.slope * a.avg_x)) as intercept, a.avg_x, a.avg_y, s.slope
into #intercept
from #avg as a join #slope s on a.site = s.site and a.fymonth = s.fymonth
order by a.site, a.fymonth


-- calculate forecast
select distinct 
'Forecast: Annual Giving ' source,
c.fiscalyear, 
c.firstofmonth,
c.fy_month_num,
i.site, 
i.sitename,
i.sitegroup,
'Annual Giving' gift_level,
'' gift_type,
c.fiscalyear * i.slope + i.intercept amount

into
#agforecast

from 
calendar c, #intercept i

where 
c.fiscalyear between 2024 and 2028 
and c.firstofmonth > datefromparts(year(getdate()), month(getdate()), 1)
and c.fy_month_num = i.fymonth


-------------------------------------------------------------------------------
-- Imported EDFD and WUNC data
-------------------------------------------------------------------------------
select *
into #imported
from
(
select
seq_asc seq,
fiscalyear, 
datefromparts(year(asofdate), month(asofdate), 1) date,
case when month(asofdate) < 7 then month(asofdate) + 6 else  month(asofdate) - 6 end fymonth,
site, 
case when site = 'edfd' then 'Educational Foundation' else 'WUNC-FM' end sitename, 
--'' designationnumber, 
--'' designation, 
'EDFN-WUNC' gift_level, 
'Gift' gift_type, 
--commitment_currentfy_gift_amount amount,
commitment_currentfy_gift_amount amount

from
fytd_special

union

select
seq_asc seq,
fiscalyear, 
datefromparts(year(asofdate), month(asofdate), 1) date,
case when month(asofdate) < 7 then month(asofdate) + 6 else  month(asofdate) - 6 end fymonth,
site, 
case when site = 'edfd' then 'Educational Foundation' else 'WUNC-FM' end sitename, 
--'' designationnumber, 
--'' designation, 
'EDFN-WUNC' gift_level, 
'Grant' gift_type, 
--commitment_currentfy_grant_amount amount,
commitment_currentfy_grant_amount amount

from
fytd_special

union

select
seq_asc seq,
fiscalyear, 
datefromparts(year(asofdate), month(asofdate), 1) date,
case when month(asofdate) < 7 then month(asofdate) + 6 else  month(asofdate) - 6 end fymonth,
site, 
case when site = 'edfd' then 'Educational Foundation' else 'WUNC-FM' end sitename, 
--'' designationnumber, 
--'' designation, 
'EDFN-WUNC' gift_level, 
'Pledge' gift_type, 
--commitment_currentfy_pledge_amount amount,
commitment_currentfy_pledge_amount amount

from
fytd_special

union

select
seq_asc seq,
fiscalyear, 
datefromparts(year(asofdate), month(asofdate), 1) date,
case when month(asofdate) < 7 then month(asofdate) + 6 else  month(asofdate) - 6 end fymonth,
site, 
case when site = 'edfd' then 'Educational Foundation' else 'WUNC-FM' end sitename, 
--'' designationnumber, 
--'' designation, 
'EDFN-WUNC' gift_level, 
'Planned Gift' gift_type, 
--commitment_currentfy_revocable_amount + commitment_currentfy_irrevocable_amount amount,
commitment_currentfy_revocable_amount + commitment_currentfy_irrevocable_amount amount

from
fytd_special
)x;


-- get monthly amounts from running totals
select *
into #importedbymonth
from
(
select
'Commitments: Imported EDFN and WUNC' source,
a.fiscalyear, 
a.date, 
a.fymonth,
a.site, 
a.sitename, 
a.sitename sitegroup,
--a.designationnumber, 
--a.designation, 
a.gift_level, 
a.gift_type, 
--case when a.amount - b.amount < 0 then 0 else a.amount - b.amount end amount
a.amount - b.amount amount

from 
#imported a join #imported b on a.site = b.site and a.fiscalyear = b.fiscalyear and a.gift_type = b.gift_type and a.seq = b.seq + 1

where
a.site in ('edfd', 'wunc')

union

select
'Commitments: Imported EDFN and WUNC' source,
fiscalyear, 
date, 
fymonth,
site, 
sitename, 
sitename sitegroup,
--designationnumber, 
--designation, 
gift_level, 
gift_type, 
amount

from 
#imported 

where
seq = 1 and site in ('edfd', 'wunc')
)x;


-------------------------------------------------------------------------------
-- opportunity pipeline
-------------------------------------------------------------------------------

-- calculate melt
select
s.shortname site, 
o.expectedclosedate,
o.responsedate,
case when month(o.expectedclosedate) < 7 then month(o.expectedclosedate) + 6 else  month(o.expectedclosedate) - 6 end fymonth,
case when month(o.expectedclosedate) < 7 then year(o.expectedclosedate) else year(o.expectedclosedate) - 1 end as expectedclosefy,
case when month(o.responsedate) < 7 then year(o.responsedate) else year(o.responsedate) - 1 end as closefy,
case
when (case when month(o.responsedate) < 7 then year(o.responsedate) else year(o.responsedate) - 1 end) > (case when month(o.expectedclosedate) < 7 then year(o.expectedclosedate) else year(o.expectedclosedate) - 1 end) then 1 
else 0 
end as melt_ind

into 
#melt

from
opportunity o
join opportunitydesignation od on o.id = od.opportunityid
join designations d on od.designationid = d.id
join sites s on d.siteid = s.id

where
o.status = 'accepted'
and o.expectedclosedate is not null;


select site, (1.0 * sum(melt_ind)) / (1.0 * count(melt_ind)) melt_pct
into #melt_pct
from #melt
group by site;



-- pipeline
select
'Forecast: Opportunity Pipeline' source,
--case when month(o.expectedclosedate) <=6 then year(o.expectedclosedate) else year(o.expectedclosedate) - 1 end fiscalyear,
--datefromparts(year(o.expectedclosedate), month(o.expectedclosedate), 1) date,
--case when month(o.expectedclosedate) < 7 then month(o.expectedclosedate) + 6 else  month(o.expectedclosedate) - 6 end fymonth,
c.fiscalyear,
c.firstofmonth date,
case when month(o.expectedclosedate) < 7 then month(o.expectedclosedate) + 6 else  month(o.expectedclosedate) - 6 end fymonth,
od.site,
od.sitename,
case 
when od.site in ('LCCC','MED', 'MEDF') then 'UNC Health'
when od.site in ('ADMU','LTNX', 'CHPDP', 'CHAN', 'CRC', 'FPG', 'HSRC', 'IPRC', 'OVCR', 'PROV', 'RENCI', 'SHSC') then 'Units Without Development Offices'
else od.sitename end as sitegroup,
--od.designationnumber,
--od.designation,
case 
when od.amount >= @principalgiving then 'Principal Giving'
when od.amount >= @majorgiving then 'Major Giving'
else 'Annual Giving' end as gift_level,
'Forecast' as gift_type,
(1 - od.melt_pct) * sum(od.amount * ((1.0*case when o.likelihood = 'low' then .25 when o.likelihood = 'even' then .5 else .75 end))) amount

into 
#pipeline

from 
prospectplan pp
join opportunity o on o.prospectplanid = pp.id
join calendar c on o.expectedclosedate  = c.date

outer apply 
(
select top 1
od.opportunityid, designationnumber, description designation, s.shortname site, s.name sitename, od.amount, m.melt_pct,
case 
when s.shortname in ('LCCC','MED', 'MEDF') then 'UNC Health'
when s.shortname in ('ADMU','LTNX', 'CHPDP', 'CHAN', 'CRC', 'FPG', 'HSRC', 'IPRC', 'OVCR', 'PROV', 'RENCI', 'SHSC') then 'Units Without Development Offices'
else s.name end as sitegroup

from
opportunitydesignation od
join designations d on od.designationid = d.id
join sites s on d.siteid = s.id
left join #melt_pct m on s.shortname = m.site

where
od.opportunityid = o.id
) od


where 
pp.isactive = 1
and pp.prospectplantype <> 'commitment'
and o.status not in ('canceled', 'rejected')
and isnull(o.revenuecommitted, 0) = 0
and o.expectedclosedate > getdate()
and od.amount >= @majorgiving

group by
--case when month(o.expectedclosedate) <=6 then year(o.expectedclosedate) else year(o.expectedclosedate) - 1 end,
--datefromparts(year(o.expectedclosedate), month(o.expectedclosedate), 1),
c.fiscalyear,
c.firstofmonth,
case when month(o.expectedclosedate) < 7 then month(o.expectedclosedate) + 6 else  month(o.expectedclosedate) - 6 end,
od.site,
od.sitename,
case 
when od.site in ('LCCC','MED', 'MEDF') then 'UNC Health'
when od.site in ('ADMU','LTNX', 'CHPDP', 'CHAN', 'CRC', 'FPG', 'HSRC', 'IPRC', 'OVCR', 'PROV', 'RENCI', 'SHSC') then 'Units Without Development Offices'
else ' ' + od.sitename end,
--od.designationnumber,
--od.designation,
case 
when od.amount >= @principalgiving then 'Principal Giving'
when od.amount >= @majorgiving then 'Major Giving'
else 'Annual Giving' end,
od.melt_pct

select x.*, 
convert(varchar(4), cal.year) + '-' + case when cal.month_num < 10 then '0' + convert(varchar(1), cal.month_num) else convert(varchar(2), cal.month_num) end calsort,
convert(varchar(4), cal.fiscalyear) + '-' +  case when cal.fy_month_num < 10 then '0' + convert(varchar(1), cal.fy_month_num) else convert(varchar(2), cal.fy_month_num) end fysort
from
(
select * from #commitment_tx 
union
select * from #importedbymonth 
union
select * from #pipeline 
union
select * from #agforecast
) x join calendar cal on x.date = cal.date


