

-- calculate new rating
drop table if exists #rf0;
drop table if exists #rf;
drop table if exists #all_giving;
drop table if exists #site_giving;
drop table if exists #site_giving_score;
drop table if exists #giving_unc_rating;
drop table if exists #giving_donorscape_rating;
drop table if exists #giving_capacity;
drop table if exists #capacity_score;
drop table if exists #facultystaff;
drop table if exists #board_detail;
drop table if exists #board_score;
drop table if exists #last_interactions;
drop table if exists #interaction_counts;
drop table if exists #interaction_multiplier;
drop table if exists #interaction_score0;
drop table if exists #interaction_score1;
drop table if exists #interaction_score;
drop table if exists #event_detail;
drop table if exists #event_score;
drop table if exists #sfmc_score;
drop table if exists #classnotes_score;
drop table if exists #constituent_site;
drop table if exists #plannedgift_score;
drop table if exists #volunteer_score;
drop table if exists #scores;
drop table if exists #scores_relation_new;
drop table if exists #scores_relation_update
drop table if exists #scores_plus_legacy;
drop table if exists #scores_plus_househould;
drop table if exists #percentiles_total;
drop table if exists #percentiles_philanthropic;
drop table if exists #percentiles_volunteer;
drop table if exists #percentiles_communication;
drop table if exists #percentiles_experiential;
drop table if exists #percentiles;
drop table if exists #records;

-------------------------------------------------------------------------------------------
-- philanthropic 3) recency + frequency by site rfm - 8 point maximum
-- grad 12 point maximum
-- law 10 point maximum
-- mpsc 10 point maximum
-- sog 5 point maximum
-------------------------------------------------------------------------------------------
select
constituentid, siteid, max(giftdate) max_giftdate, min(giftdate) as min_giftdate, year(max(giftdate)) - year(min(giftdate)) as yeardiff, count(id) as giftcnt, count(distinct calendaryear) as years_giving,
case when max(giftdate) >= dateadd(month, -30, getdate()) then 1 else 0 end as time_multiplier,

case
when max(giftdate) >= dateadd(year, -1, getdate()) then 20
when max(giftdate) < dateadd(year, -1, getdate()) and max(giftdate) >= dateadd(year, -2, getdate()) then 15
when max(giftdate) < dateadd(year, -2, getdate()) and max(giftdate) >= dateadd(year, -3, getdate()) then 10
when max(giftdate) < dateadd(year, -3, getdate()) and max(giftdate) >= dateadd(year, -4, getdate()) then 5
when max(giftdate) < dateadd(year, -4, getdate()) and max(giftdate) >= dateadd(year, -5, getdate()) then 2
when max(giftdate) < dateadd(year, -5, getdate()) then 1
end as recency,

(count(id) /  (1 + (year(max(giftdate))*1.0) - (year(min(giftdate))*1.0) ))  as freq_pct,

case
when (count(id) /  (1 + (year(max(giftdate))*1.0) - (year(min(giftdate))*1.0) ))  >= 1 then 10
when (count(id) /  (1 + (year(max(giftdate))*1.0) - (year(min(giftdate))*1.0) ))  >= .9 then 8
when (count(id) /  (1 + (year(max(giftdate))*1.0) - (year(min(giftdate))*1.0) ))  >= .8 then 6
when (count(id) /  (1 + (year(max(giftdate))*1.0) - (year(min(giftdate))*1.0) ))  >= .7 then 4
when (count(id) /  (1 + (year(max(giftdate))*1.0) - (year(min(giftdate))*1.0) ))  >= .6 then 2
else 1 end as freq_pct_points,

case 
when count(distinct calendaryear) >= 20 then 3
when count(distinct calendaryear) >= 15 then 2.5
when count(distinct calendaryear) >= 10 then 2
when count(distinct calendaryear) >= 5 then 1.5 
else 1 
end as years_giving_points

into
#rf0

from 
usr_unc_constituentgivingaggregatebase

where 
isedfoundation = 0 and ispledge = 0 and isplannedgift = 0 and isgrant = 0 and isrecognized = 1

group by
constituentid, siteid;


select *, 
freq_pct_points * years_giving_points as freq, 
freq_pct_points * years_giving_points + recency as rf,

case
-- grad
when siteid = 'F8A77A23-6EB1-4F36-B7A2-8E7C19C2B128' and freq_pct_points * years_giving_points + recency >= 40 then 1.0 * 12
when siteid = 'F8A77A23-6EB1-4F36-B7A2-8E7C19C2B128' and freq_pct_points * years_giving_points + recency >= 30 then 1.0 * 9
when siteid = 'F8A77A23-6EB1-4F36-B7A2-8E7C19C2B128' and freq_pct_points * years_giving_points + recency >= 20 then 1.0 * 6
when siteid = 'F8A77A23-6EB1-4F36-B7A2-8E7C19C2B128' and freq_pct_points * years_giving_points + recency >= 10 then 1.0 * 3
when siteid = 'F8A77A23-6EB1-4F36-B7A2-8E7C19C2B128' and freq_pct_points * years_giving_points + recency > 0   then 1.5

-- law
when siteid = '114F147B-B1F4-4B31-8AF9-8AE09C7B0C1F' and freq_pct_points * years_giving_points + recency >= 40 then 1.0 * 10
when siteid = '114F147B-B1F4-4B31-8AF9-8AE09C7B0C1F' and freq_pct_points * years_giving_points + recency >= 30 then 7.5
when siteid = '114F147B-B1F4-4B31-8AF9-8AE09C7B0C1F' and freq_pct_points * years_giving_points + recency >= 20 then 1.0 * 5
when siteid = '114F147B-B1F4-4B31-8AF9-8AE09C7B0C1F' and freq_pct_points * years_giving_points + recency >= 10 then 2.5

-- mpsc
when siteid = '03686AA6-2F31-47AC-8FE4-F988584DDECE' and freq_pct_points * years_giving_points + recency >= 40 then 1.0 * 10
when siteid = '03686AA6-2F31-47AC-8FE4-F988584DDECE' and freq_pct_points * years_giving_points + recency >= 30 then 7.5
when siteid = '03686AA6-2F31-47AC-8FE4-F988584DDECE' and freq_pct_points * years_giving_points + recency >= 20 then 1.0 * 5
when siteid = '03686AA6-2F31-47AC-8FE4-F988584DDECE' and freq_pct_points * years_giving_points + recency >= 10 then 2.5
when siteid = '03686AA6-2F31-47AC-8FE4-F988584DDECE' and freq_pct_points * years_giving_points + recency > 0   then 1.3

-- sog
when siteid = '38072215-631A-4059-BB84-CB2FB939593C' and freq_pct_points * years_giving_points + recency >= 40 then 1.0 * 5
when siteid = '38072215-631A-4059-BB84-CB2FB939593C' and freq_pct_points * years_giving_points + recency >= 30 then 3.8
when siteid = '38072215-631A-4059-BB84-CB2FB939593C' and freq_pct_points * years_giving_points + recency >= 20 then 2.5
when siteid = '38072215-631A-4059-BB84-CB2FB939593C' and freq_pct_points * years_giving_points + recency >= 10 then 1.3
when siteid = '38072215-631A-4059-BB84-CB2FB939593C' and freq_pct_points * years_giving_points + recency > 0   then 0.6

when freq_pct_points * years_giving_points + recency >= 40 then 1.0 * 8
when freq_pct_points * years_giving_points + recency >= 30 then 1.0 * 6
when freq_pct_points * years_giving_points + recency >= 20 then 1.0 * 4
when freq_pct_points * years_giving_points + recency >= 10 then 1.0 * 2
else 1 end as rf_score

into
#rf

from 
#rf0; 

-------------------------------------------------------------------------------------------
-- philanthropic 1) % overall giving to site (aka share of wallet) - 12 point maximum
-- grad 10 point maximum
-- mpsc 15 point maximum
-- nri 10 point maximum
-- sog 17 point maximum
-------------------------------------------------------------------------------------------
select
c.constituentsystemid as constituentid, 
sum(amount) as total_amt

into
#all_giving

from
bbinfinity_rpt_bbdw.bbdw.fact_usr_unc_aggregategiving_by_site a 
join bbinfinity_rpt_bbdw.bbdw.dim_constituent c on a.constituentdimid = c.constituentdimid

group by
c.constituentsystemid;


select
c.constituentsystemid as constituentid, 
s.sitesystemid as siteid,
s.shortname as site,
a.amount as site_amt,
g.total_amt,
case when g.total_amt = 0 then 0 else a.amount / g.total_amt end as site_pct

into
#site_giving

from
bbinfinity_rpt_bbdw.bbdw.fact_usr_unc_aggregategiving_by_site a 
join bbinfinity_rpt_bbdw.bbdw.dim_constituent c on a.constituentdimid = c.constituentdimid
join bbinfinity_rpt_bbdw.bbdw.dim_site s on a.sitedimid = s.sitedimid
join #all_giving g on c.constituentsystemid = g.constituentid;


select
s.constituentid, s.siteid, s.site, site_pct, r.time_multiplier as multiplier, site_amt, total_amt,

case 

when s.site = 'grad' and site_amt < 500 then 0
when s.site = 'grad' and s.site_pct >= .8 then 10 
when s.site = 'grad' and s.site_pct >= .6 then 8.13
when s.site = 'grad' and s.site_pct >= .4 then 6.04
when s.site = 'grad' and s.site_pct >= .2 then 4.17
when s.site = 'grad' and s.site_pct >= .1 then 2.08
when s.site = 'grad' and s.site_pct > 0 then 0.83

when s.site = 'mpsc' and site_amt < 500 then 0
when s.site = 'mpsc' and s.site_pct >= .8 then 15
when s.site = 'mpsc' and s.site_pct >= .6 then 12.19
when s.site = 'mpsc' and s.site_pct >= .4 then 9.06
when s.site = 'mpsc' and s.site_pct >= .2 then 6.25
when s.site = 'mpsc' and s.site_pct >= .1 then 3.13
when s.site = 'mpsc' and s.site_pct > 0 then 1.25

when s.site = 'nri' and site_amt < 500 then 0
when s.site = 'nri' and s.site_pct >= .8 then 10
when s.site = 'nri' and s.site_pct >= .6 then 8.13
when s.site = 'nri' and s.site_pct >= .4 then 6.04
when s.site = 'nri' and s.site_pct >= .2 then 4.17
when s.site = 'nri' and s.site_pct >= .1 then 2.08
when s.site = 'nri' and s.site_pct > 0 then 0.83

when s.site = 'sog' and site_amt < 500 then 0
when s.site = 'sog' and s.site_pct >= .8 then 17
when s.site = 'sog' and s.site_pct >= .6 then 13.81
when s.site = 'sog' and s.site_pct >= .4 then 10.27
when s.site = 'sog' and s.site_pct >= .2 then 7.08
when s.site = 'sog' and s.site_pct >= .1 then 3.54
when s.site = 'sog' and s.site_pct > 0 then 1.42

when site_amt < 500 then 0
when s.site_pct >= .8 then 12 
when s.site_pct >= .6 then 9.75 
when s.site_pct >= .4 then 7.25 
when s.site_pct >= .2 then 5 
when s.site_pct >= .1 then 2.5
when s.site_pct > 0 then 1 
else 0 end as site_giving_score

into
#site_giving_score

from
#site_giving s 
join #rf r on s.constituentid = r.constituentid and s.siteid = r.siteid and r.time_multiplier = 1


-------------------------------------------------------------------------------------------
-- philanthropic 2) lifetime giving by capacity rating by site - 10 point maximum
-- grad 5 point maximum
-- nri 6 point maximum
-- sog 8 point maximum
-------------------------------------------------------------------------------------------
select 
p.id constituentid, 
max(case
when psc.description like 'A__-%' then 1000000 -- 5000000 capped at 1,000,000 per justin
when psc.description like 'AA__-%' then 1000000 -- 10000000 capped at 1,000,000 per justin
when psc.description like 'AAA__-%' then 1000000 -- 25000000 capped at 1,000,000 per justin
when psc.description like 'AAAA__-%' then 1000000 -- 50000000 capped at 1,000,000 per justin
when psc.description like 'AAAAA__-%' then 1000000 -- 100000000 capped at 1,000,000 per justin
when psc.description like 'B__-%' then 1000000
when psc.description like 'C__-%' then 500000
when psc.description like 'D__-%' then 100000
when psc.description like 'E__-%' then 25000
when psc.description like 'XAAAAA%' then 1000000 -- 100000000 capped at 1,000,000 per justin
when psc.description like 'XAAAA%' then 1000000 -- 50000000 capped at 1,000,000 per justin
when psc.description like 'XAAA%' then 1000000 -- 25000000 capped at 1,000,000 per justin
when psc.description like 'XAA%' then 1000000 -- 10000000 capped at 1,000,000 per justin
when psc.description like 'XA%' then 1000000 -- 5000000 capped at 1,000,000 per justin
when psc.description like 'XB%' then 1000000
when psc.description like 'XC%' then 500000
when psc.description like 'XD%' then 100000
when psc.description like 'XE%' then 25000
else 0 end) as unc_rating_amt

into
#giving_unc_rating

from 
prospect p, prospectstatuscode psc

where 
p.prospectstatuscodeid = psc.id

group by
p.id;


select
constituentid, 
max(case
when ExactNearGiftCapacityRating like '1%' then 1000000 -- 10000000 capped at 1,000,000 per justin
when ExactNearGiftCapacityRating like '2%' then 1000000
when ExactNearGiftCapacityRating like '3%' then 250000
when ExactNearGiftCapacityRating like '4%' then 100000
when ExactNearGiftCapacityRating like '5%' then 25000
when ExactNearGiftCapacityRating like '6%' then 10000
when ExactNearGiftCapacityRating like '7%' then 2500
else 0 end) as donorscape_capacity_amt

into
#giving_donorscape_rating

from 
donorscape2019

group by
constituentid;


select
g.constituentid, 
g.siteid,
g.site,
g.site_amt, 
u.unc_rating_amt,
d.donorscape_capacity_amt,
case when u.unc_rating_amt is null then d.donorscape_capacity_amt else u.unc_rating_amt end as denominator

into
#giving_capacity

from 
#site_giving g
left join #giving_unc_rating u on g.constituentid = u.constituentid
left join #giving_donorscape_rating d on g.constituentid = d.constituentid;


select
g.constituentid, g.siteid, g.site, g.site_amt, r.time_multiplier as multiplier, g.unc_rating_amt, g.donorscape_capacity_amt, g.denominator,  
case when g.denominator = 0 then 0 else g.site_amt / g.denominator end calc,

case 
when g.denominator = 0 then 0

when g.site = 'grad' and g.site_amt / g.denominator >= 1 then 5.0
when g.site = 'grad' and g.site_amt / g.denominator >= .95 then 4.8
when g.site = 'grad' and g.site_amt / g.denominator >= .9 then 4.5
when g.site = 'grad' and g.site_amt / g.denominator >= .85 then 4.3
when g.site = 'grad' and g.site_amt / g.denominator >= .8 then 4.0
when g.site = 'grad' and g.site_amt / g.denominator >= .75 then 3.8
when g.site = 'grad' and g.site_amt / g.denominator >= .7 then 3.5
when g.site = 'grad' and g.site_amt / g.denominator >= .65 then 3.3
when g.site = 'grad' and g.site_amt / g.denominator >= .6 then 3.0
when g.site = 'grad' and g.site_amt / g.denominator >= .55 then 2.8
when g.site = 'grad' and g.site_amt / g.denominator >= .5 then 2.5
when g.site = 'grad' and g.site_amt / g.denominator >= .45 then 2.3
when g.site = 'grad' and g.site_amt / g.denominator >= .4 then 2.0
when g.site = 'grad' and g.site_amt / g.denominator >= .35 then 1.8
when g.site = 'grad' and g.site_amt / g.denominator >= .3 then 1.5
when g.site = 'grad' and g.site_amt / g.denominator >= .25 then 1.3
when g.site = 'grad' and g.site_amt / g.denominator >= .2 then 1.0
when g.site = 'grad' and g.site_amt / g.denominator >= .15 then 0.8
when g.site = 'grad' and g.site_amt / g.denominator >= .1 then 0.5
when g.site = 'grad' and g.site_amt / g.denominator > 0 then 0.3

when g.site = 'nri' and g.site_amt / g.denominator >= 1 then 6.0
when g.site = 'nri' and g.site_amt / g.denominator >= .95 then 5.7
when g.site = 'nri' and g.site_amt / g.denominator >= .9 then 5.4
when g.site = 'nri' and g.site_amt / g.denominator >= .85 then 5.1
when g.site = 'nri' and g.site_amt / g.denominator >= .8 then 4.8
when g.site = 'nri' and g.site_amt / g.denominator >= .75 then 4.5
when g.site = 'nri' and g.site_amt / g.denominator >= .7 then 4.2
when g.site = 'nri' and g.site_amt / g.denominator >= .65 then 3.9
when g.site = 'nri' and g.site_amt / g.denominator >= .6 then 3.6
when g.site = 'nri' and g.site_amt / g.denominator >= .55 then 3.3
when g.site = 'nri' and g.site_amt / g.denominator >= .5 then 3.0
when g.site = 'nri' and g.site_amt / g.denominator >= .45 then 2.7
when g.site = 'nri' and g.site_amt / g.denominator >= .4 then 2.4
when g.site = 'nri' and g.site_amt / g.denominator >= .35 then 2.1
when g.site = 'nri' and g.site_amt / g.denominator >= .3 then 1.8
when g.site = 'nri' and g.site_amt / g.denominator >= .25 then 1.5
when g.site = 'nri' and g.site_amt / g.denominator >= .2 then 1.2
when g.site = 'nri' and g.site_amt / g.denominator >= .15 then 0.9
when g.site = 'nri' and g.site_amt / g.denominator >= .1 then 0.6
when g.site = 'nri' and g.site_amt / g.denominator > 0 then 0.3


when g.site = 'sog' and g.site_amt / g.denominator >= 1 then 8.0
when g.site = 'sog' and g.site_amt / g.denominator >= .95 then 7.6
when g.site = 'sog' and g.site_amt / g.denominator >= .9 then 7.2
when g.site = 'sog' and g.site_amt / g.denominator >= .85 then 6.8
when g.site = 'sog' and g.site_amt / g.denominator >= .8 then 6.4
when g.site = 'sog' and g.site_amt / g.denominator >= .75 then 6.0
when g.site = 'sog' and g.site_amt / g.denominator >= .7 then 5.6
when g.site = 'sog' and g.site_amt / g.denominator >= .65 then 5.2
when g.site = 'sog' and g.site_amt / g.denominator >= .6 then 4.8
when g.site = 'sog' and g.site_amt / g.denominator >= .55 then 4.4
when g.site = 'sog' and g.site_amt / g.denominator >= .5 then 4.0
when g.site = 'sog' and g.site_amt / g.denominator >= .45 then 3.6
when g.site = 'sog' and g.site_amt / g.denominator >= .4 then 3.2
when g.site = 'sog' and g.site_amt / g.denominator >= .35 then 2.8
when g.site = 'sog' and g.site_amt / g.denominator >= .3 then 2.4
when g.site = 'sog' and g.site_amt / g.denominator >= .25 then 2.0
when g.site = 'sog' and g.site_amt / g.denominator >= .2 then 1.6
when g.site = 'sog' and g.site_amt / g.denominator >= .15 then 1.2
when g.site = 'sog' and g.site_amt / g.denominator >= .1 then 0.8
when g.site = 'sog' and g.site_amt / g.denominator > 0 then 0.4


when g.site_amt / g.denominator >= 1 then 10
when g.site_amt / g.denominator >= .95 then 9.5
when g.site_amt / g.denominator >= .9 then 9
when g.site_amt / g.denominator >= .85 then 8.5
when g.site_amt / g.denominator >= .8 then 8
when g.site_amt / g.denominator >= .75 then 7.5
when g.site_amt / g.denominator >= .7 then 7
when g.site_amt / g.denominator >= .65 then 6.5
when g.site_amt / g.denominator >= .6 then 6
when g.site_amt / g.denominator >= .55 then 5.5
when g.site_amt / g.denominator >= .5 then 5
when g.site_amt / g.denominator >= .45 then 4.5
when g.site_amt / g.denominator >= .4 then 4
when g.site_amt / g.denominator >= .35 then 3.5
when g.site_amt / g.denominator >= .3 then 3
when g.site_amt / g.denominator >= .25 then 2.5
when g.site_amt / g.denominator >= .2 then 2
when g.site_amt / g.denominator >= .15 then 1.5
when g.site_amt / g.denominator >= .1 then 1
when g.site_amt / g.denominator > 0 then .5

else 0 end as giving_score_raw,

case 
when g.denominator = 0 then 0
when g.site_amt / g.denominator >= 1 then 10 * r.time_multiplier
when g.site_amt / g.denominator >= .95 then 9.5 * r.time_multiplier
when g.site_amt / g.denominator >= .9 then 9 * r.time_multiplier
when g.site_amt / g.denominator >= .85 then 8.5 * r.time_multiplier
when g.site_amt / g.denominator >= .8 then 8 * r.time_multiplier
when g.site_amt / g.denominator >= .75 then 7.5 * r.time_multiplier
when g.site_amt / g.denominator >= .7 then 7 * r.time_multiplier
when g.site_amt / g.denominator >= .65 then 6.5 * r.time_multiplier
when g.site_amt / g.denominator >= .6 then 6 * r.time_multiplier
when g.site_amt / g.denominator >= .55 then 5.5 * r.time_multiplier
when g.site_amt / g.denominator >= .5 then 5 * r.time_multiplier
when g.site_amt / g.denominator >= .45 then 4.5 * r.time_multiplier
when g.site_amt / g.denominator >= .4 then 4 * r.time_multiplier
when g.site_amt / g.denominator >= .35 then 3.5 * r.time_multiplier
when g.site_amt / g.denominator >= .3 then 3 * r.time_multiplier
when g.site_amt / g.denominator >= .25 then 2.5 * r.time_multiplier
when g.site_amt / g.denominator >= .2 then 2 * r.time_multiplier
when g.site_amt / g.denominator >= .15 then 1.5 * r.time_multiplier
when g.site_amt / g.denominator >= .1 then 1 * r.time_multiplier
when g.site_amt / g.denominator > 0 then .5 * r.time_multiplier
else 0 end as capacity_score

into
#capacity_score

from 
#giving_capacity g
left join #rf r on g.constituentid = r.constituentid and g.siteid = r.siteid;


-------------------------------------------------------------------------------------------
-- philanthropic 4) planned giving - 5 point maximum
-------------------------------------------------------------------------------------------

select 
c.id constituentid, s.id siteid, s.shortname site, max(convert(int, pg.isrevocable)) isrevocable,
case when max(convert(int, pg.isrevocable)) = 1 then 2.5 else 5 end as plannedgift_score

into
#plannedgift_score

from 
usr_unc_plannedgift_view_all pg
join designation d on d.id = pg.designationid
join designationlevel dl on dl.id = d.designationlevel1id
join site s on s.id = dl.siteid
left join plannedgiftrelationship pgr on pg.plannedgiftid = pgr.plannedgiftid
left join relationship r on r.id = pgr.relationshipid
left join constituent c on c.id in (r.relationshipconstituentid, r.reciprocalconstituentid, pg.constituentid)

where 
pg.status in ('accepted', 'matured') and pg.giftdate >= dateadd(month, -30, getdate())

group by
c.id, s.id, s.shortname

-------------------------------------------------------------------------------------------
-- volunteer 1) site board participation - 15 point maximum
-- law 13 point maximum
-- mpsc 16 point maximum
-- nri 19.5 point maximum
-- sog 20 point maximum
-------------------------------------------------------------------------------------------
select distinct constituentid 
into #facultystaff
from constituency 
where constituencycodeid = '14EE48F3-7463-40C4-B828-5B9FE23C7E64' and dateto is null;


select
c1.id constituentid, c1.lookupid, c1.firstname + ' ' + c1.keyname constituent_name, c2.lookupid groupid, c2.keyname group_name, gt.name grouptype, max(isnull(dr.dateto, '9999-01-01')) dateto, cs.siteid

into
#board_detail

from
dbo.constituent c1
join dbo.groupmember gm on gm.memberid = c1.id
join dbo.constituent c2 on gm.groupid = c2.id
join dbo.groupdata gd on gm.groupid = gd.id
join dbo.grouptype gt on gt.id = gd.grouptypeid
join dbo.groupmemberdaterange dr on gm.id = dr.groupmemberid
join dbo.constituentsite cs on c2.id = cs.constituentid

where
gt.name in ('board', 'fundraising committee', 'committee', 'reunion committee')

group by
c1.id, c1.lookupid, c1.firstname + ' ' + c1.keyname, c2.lookupid, c2.keyname, gt.name, cs.siteid;


--select 
--b.constituentid, b.siteid, 
--max(case when f.constituentid is not null then 0
--when dateto >= getdate() then 15 
--else 7.5 end) as board_score

select 
b.constituentid, b.siteid, 
case 
when b.siteid = '114F147B-B1F4-4B31-8AF9-8AE09C7B0C1F' -- law
then max(case when f.constituentid is not null then 0 when dateto >= getdate() then 13.0 else 6.5 end)

when b.siteid = '03686AA6-2F31-47AC-8FE4-F988584DDECE' -- mpsc
then max(case when f.constituentid is not null then 0 when dateto >= getdate() then 16.0 else 8.0 end)

when b.siteid = 'CFAF8335-5D0F-41F9-966D-DED57D99C78C' -- nri
then max(case when f.constituentid is not null then 0 when dateto >= getdate() then 19.5 else 9.8 end)

when b.siteid = '38072215-631A-4059-BB84-CB2FB939593CC' -- sog
then max(case when f.constituentid is not null then 0 when dateto >= getdate() then 20.0 else 10.0 end)

else
max(case when f.constituentid is not null then 0 when dateto >= getdate() then 15.0 else 7.5 end)
end as board_score

into
#board_score

from 
#board_detail b left join #facultystaff f on b.constituentid = f.constituentid

group by
b.constituentid, b.siteid;

-------------------------------------------------------------------------------------------
-- volunteer 2) volunteer interaction - 15 point maximum
-- law 12 point maximum
-- mpsc 12 point maximum
-- nri 10 point maximum
-- sog 10 point maximum
-------------------------------------------------------------------------------------------
select constituentid, siteid, max(volunteer_score) volunteer_score
into #volunteer_score
from
(
select
i.constituentid, 
isnull(s.id, '886B39D3-E27B-4302-873F-8E19D283CBCF') siteid,  
i.date interactiondate,
ic.name category, 
isc.name subcategory,

case 
when s.shortname = 'law' and isc.name = 'Accompanied DO on a visit' then 12
when s.shortname = 'law' and isc.name = 'Engaged prospect for Principal Gift team' then 12
when s.shortname = 'law' and isc.name = 'Engaged students' then 8.8
when s.shortname = 'law' and isc.name = 'Hosted an event' then 8.8
when s.shortname = 'law' and isc.name = 'Made a peer solicitation' then 4.4
when s.shortname = 'law' and isc.name = 'Made an introduction' then 4.4
when s.shortname = 'law' and isc.name = 'Referred prospect to DO' then 4.4
when s.shortname = 'law' and isc.name = 'Reviewed a list' then 2.0

when s.shortname = 'mpsc' and isc.name = 'Accompanied DO on a visit' then 12
when s.shortname = 'mpsc' and isc.name = 'Engaged prospect for Principal Gift team' then 12
when s.shortname = 'mpsc' and isc.name = 'Engaged students' then 8.8
when s.shortname = 'mpsc' and isc.name = 'Hosted an event' then 8.8
when s.shortname = 'mpsc' and isc.name = 'Made a peer solicitation' then 4.4
when s.shortname = 'mpsc' and isc.name = 'Made an introduction' then 4.4
when s.shortname = 'mpsc' and isc.name = 'Referred prospect to DO' then 4.4
when s.shortname = 'mpsc' and isc.name = 'Reviewed a list' then 2.0

when s.shortname = 'nri' and isc.name = 'Accompanied DO on a visit' then 10
when s.shortname = 'nri' and isc.name = 'Engaged prospect for Principal Gift team' then 10
when s.shortname = 'nri' and isc.name = 'Engaged students' then 7.3
when s.shortname = 'nri' and isc.name = 'Hosted an event' then 7.3
when s.shortname = 'nri' and isc.name = 'Made a peer solicitation' then 3.7
when s.shortname = 'nri' and isc.name = 'Made an introduction' then 3.7
when s.shortname = 'nri' and isc.name = 'Referred prospect to DO' then 3.7
when s.shortname = 'nri' and isc.name = 'Reviewed a list' then 1.7

when s.shortname = 'sog' and isc.name = 'Accompanied DO on a visit' then 10
when s.shortname = 'sog' and isc.name = 'Engaged prospect for Principal Gift team' then 10
when s.shortname = 'sog' and isc.name = 'Engaged students' then 7.3
when s.shortname = 'sog' and isc.name = 'Hosted an event' then 7.3
when s.shortname = 'sog' and isc.name = 'Made a peer solicitation' then 3.7
when s.shortname = 'sog' and isc.name = 'Made an introduction' then 3.7
when s.shortname = 'sog' and isc.name = 'Referred prospect to DO' then 3.7
when s.shortname = 'sog' and isc.name = 'Reviewed a list' then 1.7

when isc.name = 'Accompanied DO on a visit' then 15
when isc.name = 'Engaged prospect for Principal Gift team' then 15
when isc.name = 'Engaged students' then 11
when isc.name = 'Hosted an event' then 11
when isc.name = 'Made a peer solicitation' then 5.5
when isc.name = 'Made an introduction' then 5.5
when isc.name = 'Referred prospect to DO' then 5.5
when isc.name = 'Reviewed a list' then 2.5

else 0 end as volunteer_score

from
interaction i
join interactionsubcategory isc on i.interactionsubcategoryid = isc.id
join interactioncategory ic on isc.interactioncategoryid = ic.id
left join interactionsite isite on i.id = isite.interactionid
left join site s on isite.siteid = s.id

where
ic.name = 'volunteer activity'
and i.date >= dateadd(month, -30, getdate())
) x

group by 
constituentid, siteid;

-------------------------------------------------------------------------------------------
-- communication 1) unit contact / personal visits - 10 point maximum
-- grad 4 point maximum
-- law 12 point maximum
-- mpsc 14 point maximum
-- sog 13 point maximum
-------------------------------------------------------------------------------------------
select 
i.constituentid, si.siteid, max(i.sequenceid) sequenceid

into 
#last_interactions

from 
interaction i
join interactionsite si on i.id = si.interactionid
join interactiontypecode itc on i.interactiontypecodeid = itc.id

where 
i.date >= dateadd(month, -30, getdate())
and i.status = 'completed'
and (itc.description in ('Text', 'Electronic', 'Phone', 'Email', 'Event') or itc.description like '%Personal Visit%')

group by 
i.constituentid, si.siteid;


select 
i.constituentid, 
si.siteid, 
case
when itc.description like '%Personal Visit%' then 'personal'
when itc.description = 'Event'  then 'event'
when itc.description in ('Text', 'Electronic', 'Phone', 'Email') then 'text_elec_phone'
end as interactiontype,
count(*) as interaction_cnt

into
#interaction_counts

from 
interaction i
join interactionsite si on i.id = si.interactionid
join interactiontypecode itc on i.interactiontypecodeid = itc.id
join #last_interactions li on i.constituentid = li.constituentid and si.siteid = li.siteid

where 
i.status = 'completed'
and (itc.description in ('Text', 'Electronic', 'Phone', 'Event', 'Email') or  itc.description like '%Personal Visit%')

group by 
i.constituentid, 
si.siteid, 
case
when itc.description like '%Personal Visit%' then 'personal'
when itc.description = 'Event'  then 'event'
when itc.description in ('Text', 'Electronic', 'Phone', 'Email') then 'text_elec_phone'
end;


select 
*,
case 
when interaction_cnt >= 5 then 1
when interaction_cnt >= 4 then .8
when interaction_cnt >= 3 then .6
when interaction_cnt >= 2 then .4
when interaction_cnt >= 1 then .2
end as interaction_multiplier

into 
#interaction_multiplier

from 
#interaction_counts;


select
i.constituentid, i.date interationdate, s.id as siteid, s.shortname as site,
case
when itc.description like '%Personal Visit%' then 'personal'
when itc.description = 'Event'  then 'event'
when itc.description in ('Text', 'Electronic', 'Phone', 'Email') then 'text_elec_phone'
else itc.description 
end as interactiontype,

case 

when s.shortname = 'grad' and itc.description like '%Personal Visit%' then 4.0
when s.shortname = 'grad' and itc.description = 'Event' then 3.0
when s.shortname = 'grad' and itc.description in ('Text', 'Electronic', 'Phone', 'Email') then 1.0

when s.shortname = 'law' and itc.description like '%Personal Visit%' then 12.0
when s.shortname = 'law' and itc.description = 'Event' then 9.0
when s.shortname = 'law' and itc.description in ('Text', 'Electronic', 'Phone', 'Email') then 3.0

when s.shortname = 'mpsc' and itc.description like '%Personal Visit%' then 14.0
when s.shortname = 'mpsc' and itc.description = 'Event' then 10.0
when s.shortname = 'mpsc' and itc.description in ('Text', 'Electronic', 'Phone', 'Email') then 3.5

when s.shortname = 'sog' and itc.description like '%Personal Visit%' then 13.0
when s.shortname = 'sog' and itc.description = 'Event' then 9.8
when s.shortname = 'sog' and itc.description in ('Text', 'Electronic', 'Phone', 'Email') then 3.3

when itc.description like '%Personal Visit%' then 10
when itc.description = 'Event' then 7.5 
when itc.description in ('Text', 'Electronic', 'Phone', 'Email') then 2.5
else 0 end as interaction_score_raw

into
#interaction_score0

from
#last_interactions l
join interaction i on i.constituentid = l.constituentid and i.sequenceid = l.sequenceid
join interactiontypecode itc on i.interactiontypecodeid = itc.id
join site s on l.siteid = s.id;


select 
s.*, m.interaction_cnt, m.interaction_multiplier, s.interaction_score_raw * m.interaction_multiplier as interaction_score

into
#interaction_score1

from 
#interaction_score0 s
join #interaction_multiplier m on s.constituentid = m.constituentid and s.siteid = m.siteid and s.interactiontype = m.interactiontype;

create index idx_interaction_score on #interaction_score1 (constituentid, siteid);


select 
constituentid, siteid, site, max(interaction_score) interaction_score

into
#interaction_score

from 
#interaction_score1 

group by
constituentid, siteid, site;

create index idx_interaction_score on #interaction_score (constituentid, siteid);

-------------------------------------------------------------------------------------------
-- communication 2) sfmc site email click-thrus - 7.5 point maximum
-- grad 12 point maximum
-- law 2 point maximum
-- mpsc 8 point maximum
-- nri 10 point maximum
-------------------------------------------------------------------------------------------
select 
constituentid, 
siteid, 
site, 
cnt,
case 
when site = 'grad' and cnt >= 5 then 12.0
when site = 'grad' and cnt >= 1 then 8.0
when site = 'law' and cnt >= 5 then 2.0
when site = 'law' and cnt >= 1 then 1.3
when site = 'mpsc' and cnt >= 5 then 8.0
when site = 'mpsc' and cnt >= 1 then 5.3
when site = 'nri' and cnt >= 5 then 10.0
when site = 'nri' and cnt >= 1 then 6.7
else 7.5 end as sfmc_score

into 
#sfmc_score

from
(
select
c.constituentid, s.id siteid, s.shortname site, count(distinct c.id) cnt 

from 
usr_unc_sfmc_job j
join usr_unc_sfmc_click c on j.jobid = c.jobid
join site s on (j.fromname like '%' + s.name + '%' or j.emailname like '%' + s.shortname + ' %')

where 
clickdate >= dateadd(month, -30, getdate())
and isnull(j.category,'') <> 'test send emails'
and charindex('''solicitcode''', url) = 0
and charindex('unsub_center.aspx', url) = 0 

group by
c.constituentid, s.id, s.shortname
) x

-------------------------------------------------------------------------------------------
-- communication 3) class notes - 2.5 point maximum
-- grad 0 point maximum
-- law 1 point maximum
-- mpsc 0 point maximum
-- nri 3 point maximum
-- sog 1.5 point maximum
-------------------------------------------------------------------------------------------
select distinct
aa.constituentid, s.id siteid, 
case
when s.shortname = 'grad' then 0.0
when s.shortname = 'law' then 1.0
when s.shortname = 'mpsc' then 0.0
when s.shortname = 'nri' then 3.0
when s.shortname = 'grad' then 1.5
else 2.5 end as classnotes_score

into
#classnotes_score

from 
usr_unc_alumniactivity aa
join site s on s.id = aa.siteid
left join usr_unc_alumniactivitytypecode aat on aa.alumniactivitytypecodeid = aat.id
left join usr_unc_alumniactivitysectioncode se on se.id = aa.alumniactivitysectioncodeid

where 
se.description = 'class notes'
and year(getdate()) - aa.publicationyear <= 3

-------------------------------------------------------------------------------------------
-- experiential 1) site specific event attendance - 7.5 point maximum
-- grad 12 point point maximum
-- law 9 point maximum
-- mpsc 9 point maximum
-- nri 13 point maximum
-- sog 10 point maximum
-------------------------------------------------------------------------------------------
select 
r.constituentid, s.id as siteid, s.shortname as site, count(*) event_cnt

into
#event_detail

from 
registrant r 
join event e on r.eventid = e.id 
join eventsite es on e.id = es.eventid
join site s on es.siteid = s.id

where 
r.attended = 1
and e.startdate >= dateadd(month, -30, getdate())

group by 
r.constituentid, s.id, s.shortname;


select 
constituentid, siteid, site,
case 

when site = 'grad' and event_cnt >= 4 then 12.0
when site = 'grad' and event_cnt = 3 then 8.0
when site = 'grad' and event_cnt = 2 then 5.6
when site = 'grad' and event_cnt = 1 then 4.0

when site = 'law' and event_cnt >= 4 then 9.0
when site = 'law' and event_cnt = 3 then 6.0
when site = 'law' and event_cnt = 2 then 4.2
when site = 'law' and event_cnt = 1 then 3.0

when site = 'mpsc' and event_cnt >= 4 then 9.0
when site = 'mpsc' and event_cnt = 3 then 6.0
when site = 'mpsc' and event_cnt = 2 then 4.2
when site = 'mpsc' and event_cnt = 1 then 3.0

when site = 'nri' and event_cnt >= 4 then 13.0
when site = 'nri' and event_cnt = 3 then 8.7
when site = 'nri' and event_cnt = 2 then 6.1
when site = 'nri' and event_cnt = 1 then 4.3

when site = 'sog' and event_cnt >= 4 then 10.0
when site = 'sog' and event_cnt = 3 then 6.7
when site = 'sog' and event_cnt = 2 then 4.7
when site = 'sog' and event_cnt = 1 then 3.3

when event_cnt >= 4 then 7.5
when event_cnt = 3 then 5 
when event_cnt = 2 then 3.5
when event_cnt = 1 then 2.5
else 0 end as event_score

into
#event_score

from
#event_detail;


-------------------------------------------------------------------------------------------
-- get list of applicable constituents/sites.  
-------------------------------------------------------------------------------------------
select *
into #constituent_site 
from
(
select x.constituentid, x.siteid, s.shortname as sitename
from #site_giving_score x join site s on x.siteid = s.id
where site_giving_score > 0
union
select x.constituentid, x.siteid, s.shortname as sitename 
from #capacity_score x join site s on x.siteid = s.id
where capacity_score > 0
union
select x.constituentid, x.siteid, s.shortname as sitename
from #plannedgift_score x join site s on x.siteid = s.id
where plannedgift_score > 0
union
select x.constituentid, x.siteid, s.shortname as sitename
from #board_score x join site s on x.siteid = s.id
where board_score > 0
union
select x.constituentid, x.siteid, s.shortname as sitename
from #volunteer_score x join site s on x.siteid = s.id
where volunteer_score > 0
union
select x.constituentid, x.siteid, s.shortname as sitename 
from #sfmc_score x join site s on x.siteid = s.id
where sfmc_score > 0
union
select x.constituentid, x.siteid, s.shortname as sitename
from #classnotes_score x join site s on x.siteid = s.id
where classnotes_score > 0
union
select x.constituentid, x.siteid, s.shortname as sitename
from #interaction_score x join site s on x.siteid = s.id
union
select x.constituentid, x.siteid, s.shortname as sitename 
from #event_score x join site s on x.siteid = s.id
where event_score > 0
)x;


-------------------------------------------------------------------------------------------
-- merge scores
-------------------------------------------------------------------------------------------
select distinct 
cs.constituentid, 
cs.siteid,
isnull(sgs.site_giving_score, 0) as SITE_GIVING_SCORE,
isnull(sc.capacity_score, 0) as CAPACITY_SCORE,
isnull(rf.rf_score, 0) as RF_SCORE,
isnull(pg.plannedgift_score, 0) as PLANNEDGIFT_SCORE,
isnull(bs.board_score, 0) as BOARD_SCORE,
isnull(vs.volunteer_score, 0) as VOLUNTEER_SCORE,
isnull(sf.sfmc_score, 0) as SFMC_CLICK_SCORE,
isnull(cn.classnotes_score, 0) as CLASSNOTES_SCORE,
isnull(ins.interaction_score, 0) as INTERACTION_SCORE,
isnull(es.event_score, 0) as EVENT_SCORE,

isnull(sgs.site_giving_score, 0)
+ isnull(sc.capacity_score, 0)
+ isnull(rf.rf_score, 0)
+ isnull(pg.plannedgift_score, 0)
+ isnull(bs.board_score, 0)
+ isnull(vs.volunteer_score, 0)
+ isnull(sf.sfmc_score, 0)
+ isnull(cn.classnotes_score, 0)
+ isnull(ins.interaction_score, 0)
+ isnull(es.event_score, 0) as TOTAL_SCORE

into
#scores

---- informational to support score calculations
--rf.min_giftdate as rf_min_giftdate,
--rf.max_giftdate as rf_max_giftdate,
--rf.giftcnt as rf_giftcnt,
--rf.years_giving as rf_years_giving,
--rf.yeardiff as rf_yeardiff,
--rf.recency as rf_recency,
--rf.freq_pct as rf_freq_pct,
--rf.freq_pct_points as rf_freq_pct_points,
--rf.years_giving_points as rf_years_giving_points,
--rf.rf as rf_rf,

--sgs.site_pct as site_giving_site_pct,
--sgs.multiplier as site_giving_multiplier,
--sgs.site_giving_score_raw,
--sgs.site_amt as site_giving_site_amt ,
--sgs.total_amt as site_giving_total_amt,

--sc.unc_rating_amt as capacity_unc_rating_amt,
--sc.donorscape_capacity_amt as donorscape_rating_amt,
--sc.denominator as capacity_denominator,
--sc.giving_score_raw as capacity_giving_score_raw,

--ins.interaction_score_raw,
--ins.interaction_cnt,
--ins.interaction_multiplier

from
#constituent_site cs
join constituent c on cs.constituentid = c.id
left join deceasedconstituent dc on c.id = dc.id
join site s on cs.siteid = s.id
left join #site_giving_score sgs on sgs.constituentid = cs.constituentid and sgs.siteid = cs.siteid
left join #capacity_score sc on cs.constituentid = sc.constituentid and cs.siteid = sc.siteid
left join #board_score bs on bs.constituentid = cs.constituentid  and bs.siteid = cs.siteid
left join #interaction_score ins on ins.constituentid = cs.constituentid and ins.siteid = cs.siteid
left join #event_score es on es.constituentid = cs.constituentid and es.siteid = cs.siteid
left join #rf rf on rf.constituentid = cs.constituentid and rf.siteid = cs.siteid
left join #plannedgift_score pg on pg.constituentid = cs.constituentid and pg.siteid = cs.siteid
left join #volunteer_score vs on vs.constituentid = cs.constituentid and vs.siteid = cs.siteid
left join #sfmc_score sf on sf.constituentid = cs.constituentid and sf.siteid = cs.siteid
left join #classnotes_score cn on cn.constituentid = cs.constituentid and cn.siteid = cs.siteid

where
dc.id is null
and c.isorganization = 0
and c.isinactive = 0

create index idx_scores on #scores (constituentid, siteid);

-------------------------------------------------------------------------------------------
-- experiential 2) generational legacy connection - 7.5 point maximum
-- grad 8 point maximum
-- law 9 point maximum
-- mpsc 3 point maximum
-- nri 7 point maximum
-- sog 0 point maximum
-------------------------------------------------------------------------------------------

-- no existing score
select distinct
r.reciprocalconstituentid as constituentid, 
s1.siteid, 
0 SITE_GIVING_SCORE,	
0 CAPACITY_SCORE,	
0 RF_SCORE,
0 PLANNEDGIFT_SCORE,	
0 BOARD_SCORE,
0 VOLUNTEER_SCORE,	
0 SFMC_CLICK_SCORE,	
0 CLASSNOTES_SCORE,
0 INTERACTION_SCORE,	
0 EVENT_SCORE,
case 
when s.shortname = 'grad' then 8.0
when s.shortname = 'law' then 9.0
when s.shortname = 'mpsc' then 3.0
when s.shortname = 'nri' then 7.0
when s.shortname = 'sog' then 0.0
else 7.5 end as TOTAL_SCORE,
case
when s.shortname = 'grad' then 8.0
when s.shortname = 'law' then 9.0
when s.shortname = 'mpsc' then 3.0
when s.shortname = 'nri' then 7.0
when s.shortname = 'sog' then 0.0
else 7.5 end as LEGACY_SCORE

into
#scores_relation_new

from 
relationship r
join relationshiptypecode rtc on r.relationshiptypecodeid = rtc.id 
join #scores s1 on s1.constituentid = r.relationshipconstituentid
left join #scores s2 on r.reciprocalconstituentid = s2.constituentid and s1.siteid = s2.siteid
join site s on s1.siteid = s.id

where
rtc.description in 
(
'Brother', 'Child', 'Daughter', 'Foster Child', 'Grandchild', 'Granddaughter', 'Grandson', 
'Half-brother', 'Half-sibling', 'Half-sister', 'Sibling', 'Sister', 'Step-sibling',
'Step-son', 'Step-sister', 'Step-daughter', 'Step-child', 'Step-brother', 'Son'
)
and s2.constituentid is null


-- existing score
select distinct
r.reciprocalconstituentid as constituentid, 
s1.siteid, 
s2.SITE_GIVING_SCORE,	
s2.CAPACITY_SCORE,	
s2.RF_SCORE,
s2.PLANNEDGIFT_SCORE,	
s2.BOARD_SCORE,
s2.VOLUNTEER_SCORE,	
s2.SFMC_CLICK_SCORE,	
s2.CLASSNOTES_SCORE,
s2.INTERACTION_SCORE,	
S2.EVENT_SCORE,	
case 
when s.shortname = 'grad' then S2.TOTAL_SCORE + 8.0
when s.shortname = 'law' then S2.TOTAL_SCORE + 9.0
when s.shortname = 'mpsc' then S2.TOTAL_SCORE + 3.0
when s.shortname = 'nri' then S2.TOTAL_SCORE + 7.0
when s.shortname = 'sog' then S2.TOTAL_SCORE + 0.0
else S2.TOTAL_SCORE + 7.5 end as TOTAL_SCORE,
case
when s.shortname = 'grad' then 8.0
when s.shortname = 'law' then 9.0
when s.shortname = 'mpsc' then 3.0
when s.shortname = 'nri' then 7.0
when s.shortname = 'sog' then 0.0
else 7.5 end as LEGACY_SCORE

into
#scores_relation_update

from 
relationship r
join relationshiptypecode rtc on r.relationshiptypecodeid = rtc.id 
join #scores s1 on s1.constituentid = r.relationshipconstituentid
left join #scores s2 on r.reciprocalconstituentid = s2.constituentid and s1.siteid = s2.siteid
join site s on s1.siteid = s.id

where
rtc.description in 
(
'Brother', 'Child', 'Daughter', 'Foster Child', 'Grandchild', 'Granddaughter', 'Grandson', 
'Half-brother', 'Half-sibling', 'Half-sister', 'Sibling', 'Sister', 'Step-sibling',
'Step-son', 'Step-sister', 'Step-daughter', 'Step-child', 'Step-brother', 'Son'
)
and s2.constituentid is not null


select * 
into #scores_plus_legacy
from
(
select s.*, 0 LEGACY_SCORE
from #scores s left join #scores_relation_update su on s.constituentid = su.constituentid and s.siteid = su.siteid
where su.constituentid is null

union

select * from #scores_relation_new

union

select * from #scores_relation_update
)x;

create index idx_scores_plus_legacy on #scores_plus_legacy (constituentid, siteid);

-- remove records where total score derives only from legacy component
delete from  #scores_plus_legacy where legacy_score = total_score;  

-- reduce total score by 1/2 if score derives only from philanthropic components
update #scores_plus_legacy set total_score = total_score / 2 where site_giving_score + capacity_score + rf_score + plannedgift_score = total_score;

-------------------------------------------------------------------------------------------
-- add spouse and final score
-------------------------------------------------------------------------------------------

select distinct
s1.*, isnull(s2.total_score, 0) SPOUSE_TOTAL_SCORE, 
case when isnull(s2.total_score, 0) > s1.total_score then isnull(s2.total_score, 0) else s1.total_score end HOUSEHOLD_TOTAL_SCORE,
s1.site_giving_score + s1.capacity_score + s1.rf_score + s1.plannedgift_score as subscore_philanthropic,
s1.board_score + s1.volunteer_score as subscore_volunteer,
s1.interaction_score + s1.sfmc_click_score + s1.classnotes_Score as subscore_communication,
s1.event_score + s1.legacy_score as subscore_experiential

into
#scores_plus_househould

from
#scores_plus_legacy s1
left join constituenthousehold ch1 on s1.constituentid = ch1.id
left join constituenthousehold ch2 on ch1.householdid = ch2.householdid and ch1.id <> ch2.id
left join #scores_plus_legacy s2 on ch2.id = s2.constituentid and s1.siteid = s2.siteid



-------------------------------------------------------------------------------------------
-- percentiles by site
-------------------------------------------------------------------------------------------

select distinct 
siteid,
percentile_disc(0.8) within group (order by total_score) over (partition by siteid) percentile_80,
percentile_disc(0.6) within group (order by total_score) over (partition by siteid) percentile_60,
percentile_disc(0.4) within group (order by total_score) over (partition by siteid) percentile_40,
percentile_disc(0.2) within group (order by total_score) over (partition by siteid) percentile_20,
percentile_disc(0.01) within group (order by total_score) over (partition by siteid) percentile_01

into
#percentiles_total

from 
#scores_plus_househould

where
total_score > 0;


select distinct 
siteid,
percentile_disc(0.8) within group (order by subscore_philanthropic) over (partition by siteid) percentile_80_philanthropic,
percentile_disc(0.6) within group (order by subscore_philanthropic) over (partition by siteid) percentile_60_philanthropic,
percentile_disc(0.4) within group (order by subscore_philanthropic) over (partition by siteid) percentile_40_philanthropic,
percentile_disc(0.2) within group (order by subscore_philanthropic) over (partition by siteid) percentile_20_philanthropic,
percentile_disc(0.01) within group (order by subscore_philanthropic) over (partition by siteid) percentile_01_philanthropic

into
#percentiles_philanthropic

from 
#scores_plus_househould

where
subscore_philanthropic > 0;


select distinct 
siteid,
percentile_disc(0.8) within group (order by subscore_volunteer) over (partition by siteid) percentile_80_volunteer,
percentile_disc(0.6) within group (order by subscore_volunteer) over (partition by siteid) percentile_60_volunteer,
percentile_disc(0.4) within group (order by subscore_volunteer) over (partition by siteid) percentile_40_volunteer,
percentile_disc(0.2) within group (order by subscore_volunteer) over (partition by siteid) percentile_20_volunteer,
percentile_disc(0.01) within group (order by subscore_volunteer) over (partition by siteid) percentile_01_volunteer

into
#percentiles_volunteer

from 
#scores_plus_househould

where
subscore_volunteer > 0;


select distinct 
siteid,
percentile_disc(0.8) within group (order by subscore_communication) over (partition by siteid) percentile_80_communication,
percentile_disc(0.6) within group (order by subscore_communication) over (partition by siteid) percentile_60_communication,
percentile_disc(0.4) within group (order by subscore_communication) over (partition by siteid) percentile_40_communication,
percentile_disc(0.2) within group (order by subscore_communication) over (partition by siteid) percentile_20_communication,
percentile_disc(0.01) within group (order by subscore_communication) over (partition by siteid) percentile_01_communication

into
#percentiles_communication

from 
#scores_plus_househould

where
subscore_communication > 0;



select distinct 
siteid,
percentile_disc(0.8) within group (order by subscore_experiential) over (partition by siteid) percentile_80_experiential,
percentile_disc(0.6) within group (order by subscore_experiential) over (partition by siteid) percentile_60_experiential,
percentile_disc(0.4) within group (order by subscore_experiential) over (partition by siteid) percentile_40_experiential,
percentile_disc(0.2) within group (order by subscore_experiential) over (partition by siteid) percentile_20_experiential,
percentile_disc(0.01) within group (order by subscore_experiential) over (partition by siteid) percentile_01_experiential

into
#percentiles_experiential

from 
#scores_plus_househould

where
subscore_experiential > 0;


select 
pt.siteid, pt.percentile_80, pt.percentile_60, pt.percentile_40, pt.percentile_20, pt.percentile_01,
pp.percentile_80_philanthropic, pp.percentile_60_philanthropic, pp.percentile_40_philanthropic, pp.percentile_20_philanthropic, pp.percentile_01_philanthropic,
pv.percentile_80_volunteer, pv.percentile_60_volunteer, pv.percentile_40_volunteer, pv.percentile_20_volunteer, pv.percentile_01_volunteer,
pc.percentile_80_communication, pc.percentile_60_communication, pc.percentile_40_communication, pc.percentile_20_communication, pc.percentile_01_communication,
pe.percentile_80_experiential, pe.percentile_60_experiential, pe.percentile_40_experiential, pe.percentile_20_experiential, pe.percentile_01_experiential

into
#percentiles

from 
#percentiles_total pt
left join #percentiles_philanthropic pp on pt.siteid = pp.siteid
left join #percentiles_volunteer pv on pt.siteid = pv.siteid
left join #percentiles_communication pc on pt.siteid = pc.siteid
left join #percentiles_experiential pe on pt.siteid = pe.siteid


select site.shortname site, site.name sitename, s.*,

p.percentile_80, p.percentile_60, p.percentile_40, p.percentile_20,
case 
when s.total_score >= p.percentile_80 then N'★★★★★'
when s.total_score >= p.percentile_60 then N'★★★★☆'
when s.total_score >= p.percentile_40 then N'★★★☆☆'
when s.total_score >= p.percentile_20 then N'★★☆☆☆'
when s.total_score >= p.percentile_01 then N'★☆☆☆☆'
else N'☆☆☆☆☆'
end star_rating,

p.percentile_80_philanthropic, p.percentile_60_philanthropic, p.percentile_40_philanthropic, p.percentile_20_philanthropic,
case 
when s.subscore_philanthropic >= p.percentile_80_philanthropic then N'★★★★★'
when s.subscore_philanthropic >= p.percentile_60_philanthropic then N'★★★★☆'
when s.subscore_philanthropic >= p.percentile_40_philanthropic then N'★★★☆☆'
when s.subscore_philanthropic >= p.percentile_20_philanthropic then N'★★☆☆☆'
when s.subscore_philanthropic >= p.percentile_01_philanthropic then N'★☆☆☆☆'
else N'☆☆☆☆☆'
end star_rating_philanthropic,

p.percentile_80_volunteer, p.percentile_60_volunteer, p.percentile_40_volunteer, p.percentile_20_volunteer,
case 
when s.subscore_volunteer >= p.percentile_80_volunteer then N'★★★★★'
when s.subscore_volunteer >= p.percentile_60_volunteer then N'★★★★☆'
when s.subscore_volunteer >= p.percentile_40_volunteer then N'★★★☆☆'
when s.subscore_volunteer >= p.percentile_20_volunteer then N'★★☆☆☆'
when s.subscore_volunteer >= p.percentile_01_volunteer then N'★☆☆☆☆'
else N'☆☆☆☆☆'
end star_rating_volunteer,

p.percentile_80_communication, p.percentile_60_communication, p.percentile_40_communication, p.percentile_20_communication,
case 
when s.subscore_communication >= p.percentile_80_communication then N'★★★★★'
when s.subscore_communication >= p.percentile_60_communication then N'★★★★☆'
when s.subscore_communication >= p.percentile_40_communication then N'★★★☆☆'
when s.subscore_communication >= p.percentile_20_communication then N'★★☆☆☆'
when s.subscore_communication >= p.percentile_01_communication then N'★☆☆☆☆'
else N'☆☆☆☆☆'
end star_rating_communication,

p.percentile_80_experiential, p.percentile_60_experiential, p.percentile_40_experiential, p.percentile_20_experiential,
case 
when s.subscore_experiential >= p.percentile_80_experiential then N'★★★★★'
when s.subscore_experiential >= p.percentile_60_experiential then N'★★★★☆'
when s.subscore_experiential >= p.percentile_40_experiential then N'★★★☆☆'
when s.subscore_experiential >= p.percentile_20_experiential then N'★★☆☆☆'
when s.subscore_experiential >= p.percentile_01_experiential then N'★☆☆☆☆'
else N'☆☆☆☆☆'
end star_rating_experiential

into
#records

from 
#scores_plus_househould s join #percentiles p on s.siteid = p.siteid
join site on s.siteid = site.id

where 
site.siteid not in (' ', '35', '77', '90', '99')
--s.siteid = 'FD9A4ECE-4FDD-4A6D-8309-BE34EB9C8B40'
--s.constituentid = 'b5ca812c-e9a1-497e-8891-c262f37ea192'


insert into USR_UNC_UNCSITEENGAGEMENT_SCORE_STAR
(
ID,
SITEID,
SITE,
SITENAME,
CONSTITUENTID,
SITE_GIVING_SCORE,
CAPACITY_SCORE,
RF_SCORE,
PLANNEDGIFT_SCORE,
BOARD_SCORE,
VOLUNTEER_SCORE,
SFMC_CLICK_SCORE,
CLASSNOTES_SCORE,
INTERACTION_SCORE,
EVENT_SCORE,
LEGACY_SCORE,
TOTAL_SCORE,
SPOUSE_TOTAL_SCORE,
HOUSEHOLD_TOTAL_SCORE,
SUBSCORE_PHILANTHROPIC,
SUBSCORE_VOLUNTEER,
SUBSCORE_COMMUNICATION,
SUBSCORE_EXPERIENTIAL,
PERCENTILE_80,
PERCENTILE_60,
PERCENTILE_40,
PERCENTILE_20,
PERCENTILE_80_PHILANTHROPIC,
PERCENTILE_60_PHILANTHROPIC,
PERCENTILE_40_PHILANTHROPIC,
PERCENTILE_20_PHILANTHROPIC,
PERCENTILE_80_VOLUNTEER,
PERCENTILE_60_VOLUNTEER,
PERCENTILE_40_VOLUNTEER,
PERCENTILE_20_VOLUNTEER,
PERCENTILE_80_COMMUNICATION,
PERCENTILE_60_COMMUNICATION,
PERCENTILE_40_COMMUNICATION,
PERCENTILE_20_COMMUNICATION,
PERCENTILE_80_EXPERIENTIAL,
PERCENTILE_60_EXPERIENTIAL,
PERCENTILE_40_EXPERIENTIAL,
PERCENTILE_20_EXPERIENTIAL,
STAR_RATING,
STAR_RATING_PHILANTHROPIC,
STAR_RATING_VOLUNTEER,
STAR_RATING_COMMUNICATION,
STAR_RATING_EXPERIENTIAL,
ADDEDBYID,
CHANGEDBYID,
DATEADDED,
DATECHANGED
)

select 
newid() as ID,
SITEID,
SITE,
SITENAME,
CONSTITUENTID,
SITE_GIVING_SCORE,
CAPACITY_SCORE,
RF_SCORE,
PLANNEDGIFT_SCORE,
BOARD_SCORE,
VOLUNTEER_SCORE,
SFMC_CLICK_SCORE,
CLASSNOTES_SCORE,
INTERACTION_SCORE,
EVENT_SCORE,
LEGACY_SCORE,
TOTAL_SCORE,
SPOUSE_TOTAL_SCORE,
HOUSEHOLD_TOTAL_SCORE,
SUBSCORE_PHILANTHROPIC,
SUBSCORE_VOLUNTEER,
SUBSCORE_COMMUNICATION,
SUBSCORE_EXPERIENTIAL,
isnull(PERCENTILE_80, 0),
isnull(PERCENTILE_60, 0),
isnull(PERCENTILE_40, 0) PERCENTILE_40,
isnull(PERCENTILE_20, 0),
isnull(PERCENTILE_80_PHILANTHROPIC, 0),
isnull(PERCENTILE_60_PHILANTHROPIC, 0),
isnull(PERCENTILE_40_PHILANTHROPIC, 0),
isnull(PERCENTILE_20_PHILANTHROPIC, 0),
isnull(PERCENTILE_80_VOLUNTEER, 0),
isnull(PERCENTILE_60_VOLUNTEER, 0),
isnull(PERCENTILE_40_VOLUNTEER, 0),
isnull(PERCENTILE_20_VOLUNTEER, 0),
isnull(PERCENTILE_80_COMMUNICATION, 0),
isnull(PERCENTILE_60_COMMUNICATION, 0),
isnull(PERCENTILE_40_COMMUNICATION, 0),
isnull(PERCENTILE_20_COMMUNICATION, 0),
isnull(PERCENTILE_80_EXPERIENTIAL, 0),
isnull(PERCENTILE_60_EXPERIENTIAL, 0),
isnull(PERCENTILE_40_EXPERIENTIAL, 0),
isnull(PERCENTILE_20_EXPERIENTIAL, 0),
STAR_RATING,
STAR_RATING_PHILANTHROPIC,
STAR_RATING_VOLUNTEER,
STAR_RATING_COMMUNICATION,
STAR_RATING_EXPERIENTIAL,
@CHANGEAGENTID as ADDEDBYID,
@CHANGEAGENTID as CHANGEDBYID,
@CURRENTDATE as DATEADDED,
@CURRENTDATE as DATECHANGED

from
#records