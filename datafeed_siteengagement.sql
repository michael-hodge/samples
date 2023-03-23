
-- 1) rfm - 20 point maximum
with rf0 as
(
select
constituentid, siteid, max(giftdate) max_giftdate, min(giftdate) as min_giftdate, year(max(giftdate)) - year(min(giftdate)) as yeardiff, count(id) as giftcnt, count(distinct calendaryear) as years_giving,

case
when max(giftdate) >= dateadd(month, -6, getdate()) then 1
when max(giftdate) < dateadd(month, -6, getdate()) and max(giftdate) >= dateadd(year, -1, getdate()) then .9
when max(giftdate) < dateadd(year, -1, getdate()) and max(giftdate) >= dateadd(year, -2, getdate()) then .8
when max(giftdate) < dateadd(year, -2, getdate()) and max(giftdate) >= dateadd(year, -3, getdate()) then .7
when max(giftdate) < dateadd(year, -3, getdate()) and max(giftdate) >= dateadd(year, -4, getdate()) then .6
when max(giftdate) < dateadd(year, -4, getdate()) and max(giftdate) >= dateadd(year, -5, getdate()) then .5
when max(giftdate) < dateadd(year, -5, getdate()) then .25
end as multiplier,

case when max(giftdate) >= dateadd(year, -5, getdate()) then 1 else .5
end as fifty_all_multiplier,

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

from 
usr_unc_constituentgivingaggregatebase

where 
isedfoundation = 0 and ispledge = 0 and isplannedgift = 0 and isgrant = 0 and isrecognized=1

group by
constituentid, siteid
),

rf as
(
select *, 
freq_pct_points * years_giving_points as freq, 
freq_pct_points * years_giving_points + recency as rf,

case
when freq_pct_points * years_giving_points + recency >= 40 then 20
when freq_pct_points * years_giving_points + recency >= 30 then 15
when freq_pct_points * years_giving_points + recency >= 20 then 10
when freq_pct_points * years_giving_points + recency >= 10 then 5
else 1 end as rf_score

from 
rf0 
),

-- 2) site giving - 20 point maximum
all_giving as
(
select
c.constituentsystemid as constituentid, 
sum(amount) as total_amt

from
bbinfinity_rpt_bbdw.bbdw.fact_usr_unc_aggregategiving_by_site a 
join bbinfinity_rpt_bbdw.bbdw.dim_constituent c on a.constituentdimid = c.constituentdimid

group by
c.constituentsystemid 
),

site_giving as
(
select
c.constituentsystemid as constituentid, 
s.sitesystemid as siteid,
s.shortname as site,
a.amount as site_amt,
g.total_amt,
case when g.total_amt = 0 then 0 else a.amount / g.total_amt end as site_pct

from
bbinfinity_rpt_bbdw.bbdw.fact_usr_unc_aggregategiving_by_site a 
join bbinfinity_rpt_bbdw.bbdw.dim_constituent c on a.constituentdimid = c.constituentdimid
join bbinfinity_rpt_bbdw.bbdw.dim_site s on a.sitedimid = s.sitedimid
join all_giving g on c.constituentsystemid = g.constituentid
),

site_giving_score as
(
select
s.constituentid, s.siteid, s.site, site_pct, r.fifty_all_multiplier as multiplier, site_amt, total_amt,

case 
when site_amt < 500 then 0
when s.site_pct >= .8 then 20 
when s.site_pct >= .6 then 16 
when s.site_pct >= .4 then 12 
when s.site_pct >= .2 then 8 
when s.site_pct >= .1 then 4 
when s.site_pct > 0 then 1 
else 0 end as site_giving_score_raw,

case 
when site_amt < 500 then 0
when s.site_pct >= .8 then 20 * r.fifty_all_multiplier
when s.site_pct >= .6 then 16 * r.fifty_all_multiplier
when s.site_pct >= .4 then 12 * r.fifty_all_multiplier
when s.site_pct >= .2 then 8 * r.fifty_all_multiplier
when s.site_pct >= .1 then 4 * r.fifty_all_multiplier
when s.site_pct > 0 then 1 * r.fifty_all_multiplier
else 0 end as site_giving_score

from
site_giving s 
join rf r on s.constituentid = r.constituentid and s.siteid = r.siteid
),

-- 3) site giving by capacity - 20 point maximum
giving_unc_rating as
(
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

from 
dbo.prospect p, dbo.prospectstatuscode psc

where 
p.prospectstatuscodeid = psc.id

group by
p.id
),

giving_donorscape_rating as
(
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

from 
dbo.donorscape2019

group by
constituentid
),

giving_capacity as
(
select
g.constituentid, 
g.siteid,
g.site,
g.site_amt, 
u.unc_rating_amt,
d.donorscape_capacity_amt,
case when u.unc_rating_amt is null then d.donorscape_capacity_amt else u.unc_rating_amt end as denominator

from 
site_giving g
left join giving_unc_rating u on g.constituentid = u.constituentid
left join giving_donorscape_rating d on g.constituentid = d.constituentid
),

capacity_score as
(
select
g.constituentid, g.siteid, g.site, r.fifty_all_multiplier as multiplier, g.unc_rating_amt, g.donorscape_capacity_amt, g.denominator, 

case 
when g.denominator = 0 then 0
when g.site_amt / g.denominator >= 1 then 20
when g.site_amt / g.denominator >= .95 then 19
when g.site_amt / g.denominator >= .9 then 18
when g.site_amt / g.denominator >= .85 then 17
when g.site_amt / g.denominator >= .8 then 16
when g.site_amt / g.denominator >= .75 then 15
when g.site_amt / g.denominator >= .7 then 14
when g.site_amt / g.denominator >= .65 then 13
when g.site_amt / g.denominator >= .6 then 12
when g.site_amt / g.denominator >= .55 then 11
when g.site_amt / g.denominator >= .5 then 10
when g.site_amt / g.denominator >= .45 then 9
when g.site_amt / g.denominator >= .4 then 8
when g.site_amt / g.denominator >= .35 then 7
when g.site_amt / g.denominator >= .3 then 6
when g.site_amt / g.denominator >= .25 then 5
when g.site_amt / g.denominator >= .2 then 4
when g.site_amt / g.denominator >= .15 then 3
when g.site_amt / g.denominator >= .1 then 2
when g.site_amt / g.denominator > 0 then 1
else 0 end as giving_score_raw,

case 
when g.denominator = 0 then 0
when g.site_amt / g.denominator >= 1 then 20 * r.fifty_all_multiplier
when g.site_amt / g.denominator >= .95 then 19 * r.fifty_all_multiplier
when g.site_amt / g.denominator >= .9 then 18 * r.fifty_all_multiplier
when g.site_amt / g.denominator >= .85 then 17 * r.fifty_all_multiplier
when g.site_amt / g.denominator >= .8 then 16 * r.fifty_all_multiplier
when g.site_amt / g.denominator >= .75 then 15 * r.fifty_all_multiplier
when g.site_amt / g.denominator >= .7 then 14 * r.fifty_all_multiplier
when g.site_amt / g.denominator >= .65 then 13 * r.fifty_all_multiplier
when g.site_amt / g.denominator >= .6 then 12 * r.fifty_all_multiplier
when g.site_amt / g.denominator >= .55 then 11 * r.fifty_all_multiplier
when g.site_amt / g.denominator >= .5 then 10 * r.fifty_all_multiplier
when g.site_amt / g.denominator >= .45 then 9 * r.fifty_all_multiplier
when g.site_amt / g.denominator >= .4 then 8 * r.fifty_all_multiplier
when g.site_amt / g.denominator >= .35 then 7 * r.fifty_all_multiplier
when g.site_amt / g.denominator >= .3 then 6 * r.fifty_all_multiplier
when g.site_amt / g.denominator >= .25 then 5 * r.fifty_all_multiplier
when g.site_amt / g.denominator >= .2 then 4 * r.fifty_all_multiplier
when g.site_amt / g.denominator >= .15 then 3 * r.fifty_all_multiplier
when g.site_amt / g.denominator >= .1 then 2 * r.fifty_all_multiplier
when g.site_amt / g.denominator > 0 then 1 * r.fifty_all_multiplier
else 0 end as capacity_score

from 
giving_capacity g
left join rf r on g.constituentid = r.constituentid and g.siteid = r.siteid
),

-- 4) unc boards - 15 point maximum
facultystaff as
(
select distinct constituentid 
from constituency 
where constituencycodeid = '14EE48F3-7463-40C4-B828-5B9FE23C7E64' and dateto is null
),

board_detail as
(
select
c1.id constituentid, c1.lookupid, c1.firstname + ' ' + c1.keyname constituent_name, c2.lookupid groupid, c2.keyname group_name, gt.name grouptype, max(isnull(dr.dateto, '9999-01-01')) dateto, cs.siteid

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
c1.id, c1.lookupid, c1.firstname + ' ' + c1.keyname, c2.lookupid, c2.keyname, gt.name, cs.siteid 
),

board_score as
(
select 
b.constituentid, b.siteid, 
max(case when f.constituentid is not null then 0
when dateto >= getdate() then 15 
else 7.5 end) as board_score

from 
board_detail b left join facultystaff f on b.constituentid = f.constituentid

group by
b.constituentid, b.siteid
),

-- 5) interactions - 10 point maximum
last_interactions as 
(
select 
i.constituentid, si.siteid, max(i.sequenceid) sequenceid

from 
interaction i
join interactionsite si on i.id = si.interactionid

group by 
i.constituentid, si.siteid
),

interaction_counts as
(
select 
i.constituentid, 
si.siteid, 
case
when itc.description like '%Personal Visit%' then 'personal'
when itc.description = 'Event'  then 'event'
when itc.description in ('Text', 'Electronic', 'Phone') then 'text_elec_phone'
end as interactiontype,
count(*) as interaction_cnt

from 
interaction i
join interactionsite si on i.id = si.interactionid
join interactiontypecode itc on i.interactiontypecodeid = itc.id

where 
i.status = 'completed'

group by 
i.constituentid, 
si.siteid, 
case
when itc.description like '%Personal Visit%' then 'personal'
when itc.description = 'Event'  then 'event'
when itc.description in ('Text', 'Electronic', 'Phone') then 'text_elec_phone'
end
),

interaction_multiplier as
(
select 
*,
case 
when interaction_cnt >= 5 then 1
when interaction_cnt >= 4 then .8
when interaction_cnt >= 3 then .6
when interaction_cnt >= 2 then .4
when interaction_cnt >= 1 then .2
end as interaction_multiplier

from 
interaction_counts
),

interaction_score0 as
(
select
i.constituentid, i.date interationdate, s.id as siteid, s.shortname as site,
case
when itc.description like '%Personal Visit%' then 'personal'
when itc.description = 'Event'  then 'event'
when itc.description in ('Text', 'Electronic', 'Phone') then 'text_elec_phone'
else itc.description 
end as interactiontype,

case 
when itc.description like '%Personal Visit%' then 10 
when itc.description = 'Event' then 7.5 
when itc.description in ('Text', 'Electronic', 'Phone') then 2.5
else 0 end as interaction_score_raw

from
last_interactions l
join interaction i on i.constituentid = l.constituentid and i.sequenceid = l.sequenceid
join interactiontypecode itc on i.interactiontypecodeid = itc.id
join site s on l.siteid = s.id
), 

interaction_score as
(
select 
s.*, m.interaction_cnt, m.interaction_multiplier, s.interaction_score_raw * m.interaction_multiplier as interaction_score

from 
interaction_score0 s
join interaction_multiplier m on s.constituentid = m.constituentid and s.siteid = m.siteid and s.interactiontype = m.interactiontype
),

-- 6) event attendance - 5 point maximum
event_detail as
(
select 
r.constituentid, s.id as siteid, s.shortname as site, count(*) event_cnt, max(e.startdate) as last_event_date,

case
when max(e.startdate) >= dateadd(month, -6, getdate()) then 1
when max(e.startdate) < dateadd(month, -6, getdate()) and max(e.startdate)  >= dateadd(year, -1, getdate()) then .9
when max(e.startdate) < dateadd(year, -1, getdate()) and max(e.startdate)  >= dateadd(year, -2, getdate()) then .8
when max(e.startdate) < dateadd(year, -2, getdate()) and max(e.startdate)  >= dateadd(year, -3, getdate()) then .7
when max(e.startdate) < dateadd(year, -3, getdate()) and max(e.startdate)  >= dateadd(year, -4, getdate()) then .6
when max(e.startdate) < dateadd(year, -4, getdate()) and max(e.startdate)  >= dateadd(year, -5, getdate()) then .5
when max(e.startdate) < dateadd(year, -5, getdate()) then .25
end as recency

from 
registrant r 
join event e on r.eventid = e.id 
join eventsite es on e.id = es.eventid
join site s on es.siteid = s.id

where 
r.attended = 1

group by 
r.constituentid, s.id, s.shortname
),

event_score as
(
select 
constituentid, siteid, site, last_event_date, recency,

case 
when event_cnt >= 4 then 5
when event_cnt = 3 then 4 
when event_cnt = 2 then 3
when event_cnt = 1 then 2
else 0 end as event_score_raw,

case 
when event_cnt >= 4 then 5 * recency
when event_cnt = 3 then 4  * recency
when event_cnt = 2 then 3 * recency
when event_cnt = 1 then 2 * recency
else 0 end as event_score

from
event_detail
),

-- 7) generational legacy - 5 point maximum
unc as
(
select distinct
eh.constituentid, ea.description as college,
case 
when ea.description  = 'Arts and Sciences' then 'CAS'
when ea.description  = 'Dentistry' then 'DENT'
when ea.description  = 'Media and Journalism' then 'JOMC'
when ea.description  = 'Business' then 'KFBS'
when ea.description  = 'Law' then 'LAW'
when ea.description  = 'Medicine' then 'MED'
when ea.description  = 'Pharmacy' then 'PHAR'
when ea.description  = 'Information and Library Science' then 'SILS'
when ea.description  = 'Education' then 'SOE'
when ea.description  = 'Government' then 'SOG'
when ea.description  = 'Nursing' then 'SON'
when ea.description  = 'Public Health' then 'SPH'
when ea.description  = 'Social Work' then 'SSW'
else null end as site,

case 
when ea.description  = 'Arts and Sciences' then 'FD9A4ECE-4FDD-4A6D-8309-BE34EB9C8B40'
when ea.description  = 'Dentistry' then '109B9FA6-D198-40A3-A2E2-12D571E175CD'
when ea.description  = 'Media and Journalism' then 'E718321A-6005-4DD0-8BF1-32D704D3FBBF'
when ea.description  = 'Business' then 'F72DF0BA-E632-4E1C-845D-F8FF8A369499'
when ea.description  = 'LAW' then '114F147B-B1F4-4B31-8AF9-8AE09C7B0C1F'  
when ea.description  = 'Medicine' then 'C5033DBB-6508-4265-B3AA-8F956BF0DF94'
when ea.description  = 'Pharmacy' then 'C16B18A6-841F-412E-B264-3EE516198FBA'
when ea.description  = 'Information and Library Science' then 'B7410572-4BD6-4EDD-8B14-1F1446C7896F'
when ea.description  = 'Education' then 'EBE1F3E9-5D84-478F-9840-8B1010CBD2A7'
when ea.description  = 'Government' then '38072215-631A-4059-BB84-CB2FB939593C'
when ea.description  = 'Nursing' then '703450A9-635A-4938-B4DA-2793E80D3A5E'
when ea.description  = 'Public Health' then 'A7B5A880-91AD-4B39-81CB-B301A51BBF95'
when ea.description  = 'Social Work' then 'B3D7F9ED-3533-4A5D-946A-38D40C3676F7'
else null end as siteid

from 
educationalhistory eh 
join educationaldegreecode edc on eh.educationaldegreecodeid = edc.id 
join educationadditionalinformation ai on eh.id = ai.educationalhistoryid
join educationalcollegecode ea on ai.educationalcollegecodeid = ea.id

where 
eh.educationalinstitutionid = '67673e25-caf3-4d45-a4d2-ef633736b7d0'
),

legacy_score as
(
select distinct
r.reciprocalconstituentid as constituentid, u.siteid, u.site, 5 as legacy_score

from 
relationship r
join relationshiptypecode rtc on r.relationshiptypecodeid = rtc.id 
join constituent c1 on r.relationshipconstituentid = c1.id
join constituent c2 on r.reciprocalconstituentid = c2.id
join unc u on c1.id = u.constituentid

where
rtc.description in 
(
'Brother', 'Child', 'Daughter', 'Foster Child', 'Grandchild', 'Granddaughter', 'Grandson', 
'Half-brother', 'Half-sibling', 'Half-sister', 'Sibling', 'Sister', 'Step-sibling',
'Step-son', 'Step-sister', 'Step-daughter', 'Step-child', 'Step-brother', 'Son')
),

-- get list of applicable constituents/sites.  
constituent_site as
(
select x.constituentid, x.siteid, s.shortname as sitename
from site_giving_score x join site s on x.siteid = s.id
union
select x.constituentid, x.siteid, s.shortname as sitename 
from capacity_score x join site s on x.siteid = s.id
union
select x.constituentid, x.siteid, s.shortname as sitename
from interaction_score x join site s on x.siteid = s.id
union
select x.constituentid, x.siteid, s.shortname as sitename 
from event_score x join site s on x.siteid = s.id
union
select x.constituentid, x.siteid, s.shortname as sitename 
from legacy_score x join site s on x.siteid = s.id
),

-- merge scores
scores as
(
select distinct 
d.CONSTITUENTSYSTEMID,
d.CONSTITUENTLOOKUPID,
d.CONSTITUENTDIMID,
d.CONSTITUENTLASTNAME,
d.CONSTITUENTFULLNAME,
d.PRIMARYCONSTITUENCY,
d.GENDER,
d.ISALUMNUS,
d.ISINACTIVE,
d.ISORGANIZATION,
d.SPOUSECONSTITUENTSYSTEMID,
d.SPOUSECONSTITUENTLOOKUPID,
d.SPOUSEFULLNAME,
d.HOUSEHOLDSYSTEMID,
d.HOUSEHOLDLOOKUPID,
s.id as SITEID,
s.shortname as SITE,
isnull(sgs.site_giving_score, 0) as SITE_GIVING_SCORE,
isnull(cs.capacity_score, 0) as CAPACITY_SCORE,
isnull(bs.board_score, 0) as BOARD_SCORE,
isnull(ins.interaction_score, 0) as INTERACTION_SCORE,
isnull(es.event_score, 0) as EVENT_SCORE,
isnull(ls.legacy_score, 0) as LEGACY_SCORE,
isnull(rf.rf_score, 0) as RF_SCORE,

isnull(sgs.site_giving_score, 0) 
+ isnull(cs.capacity_score, 0) 
+ isnull(bs.board_score, 0)
+ isnull(ins.interaction_score, 0)
+ isnull(es.event_score, 0)
+ isnull(ls.legacy_score, 0)
+ isnull(rf.rf_score, 0) as TOTAL_SCORE,

0 as SPOUSE_TOTAL_SCORE,

isnull(sgs.site_giving_score, 0) 
+ isnull(cs.capacity_score, 0) 
+ isnull(bs.board_score, 0)
+ isnull(ins.interaction_score, 0)
+ isnull(es.event_score, 0)
+ isnull(ls.legacy_score, 0)
+ isnull(rf.rf_score, 0) as FINAL_TOTAL_SCORE,

0 as POINT_CHANGE,
null as PLUS_MINUS,

getdate() as CREATED_DATE,
'Current' as RECORD_IND


from
bbinfinity_rpt_bbdw.bbdw.fact_usr_unc_demographic d 
join constituent_site csite on csite.constituentid = d.constituentsystemid
join site s on csite.siteid = s.id
left join site_giving_score sgs on sgs.constituentid = csite.constituentid and sgs.siteid = csite.siteid
left join capacity_score cs on cs.constituentid = csite.constituentid and cs.siteid = csite.siteid
left join board_score bs on bs.constituentid = csite.constituentid  and bs.siteid = csite.siteid
left join interaction_score ins on ins.constituentid = csite.constituentid and ins.siteid = csite.siteid
left join event_score es on es.constituentid = csite.constituentid and es.siteid = csite.siteid
left join legacy_score ls on ls.constituentid = csite.constituentid and ls.siteid = csite.siteid
left join rf on rf.constituentid = csite.constituentid and rf.siteid = csite.siteid

where
d.isconstituent = 1
and d.isorganization = 0 
and d.isgroup = 0
and d.isinactive = 0
)

select * from scores
