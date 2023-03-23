

-- if no entry in aggregate_update_records with reset_in = 1, then only update changed records.
-- otherwise update all records
declare @reset_ind int = 
(select max(reset_ind) from bbinfinity_rpt_bbdw.dbo.aggregate_update_records);


-- anchor records to get household amount when individual amount doesn't exist for a group by dimension.
with ids as
(
select distinct ch.householdid, c.constituentsystemid constituentid, c.constituentdimid
from constituenthousehold ch join bbinfinity_rpt_bbdw.bbdw.dim_constituent c on ch.id = c.constituentsystemid
where c.constituentdimid in (select constituentdimid from bbinfinity_rpt_bbdw.dbo.aggregate_update_records) or @reset_ind = 1
),

dimensions as
(
select distinct ch.householdid, a.siteid
from usr_unc_constituentgivingaggregatebase a 
join constituenthousehold ch on a.constituentid = ch.id 
where a.isedfoundation=0 and a.iscampaign=1 and a.isrecognized=1
),

anchor as
(
select d.householdid, i.constituentid, i.constituentdimid, d.siteid
from dimensions d join ids i on d.householdid = i.householdid
),

-- household total
household_amt as
(
select
ch.householdid, 
a.siteid,
sum(a.amount) amount

from usr_unc_constituentgivingaggregatebase a
inner join bbinfinity_rpt_bbdw.bbdw.dim_constituent b on a.constituentid = b.constituentsystemid
inner join constituenthousehold ch on a.constituentid = ch.id

where 
a.isedfoundation=0 and a.iscampaign=1 and a.isrecognized=1 and a.isselfrecognition = 1
and (a.constituentid in (select constituentid from bbinfinity_rpt_bbdw.dbo.aggregate_update_records) or @reset_ind = 1)

group by
ch.householdid, a.siteid
),

-- constituent total
constituent_amt as
(
select
ch.householdid,
b.constituentdimid, 
a.siteid,
sum(a.amount) amount

from usr_unc_constituentgivingaggregatebase a
inner join bbinfinity_rpt_bbdw.bbdw.dim_constituent b on a.constituentid = b.constituentsystemid
left join constituenthousehold ch on a.constituentid = ch.id

where 
a.isedfoundation=0 and a.iscampaign=1 and a.isrecognized=1
and (a.constituentid in (select constituentid from bbinfinity_rpt_bbdw.dbo.aggregate_update_records) or @reset_ind = 1)

group by
b.constituentdimid, ch.householdid, a.siteid
)

-- merge
select
a.constituentdimid constituentdimid,
s.sitedimid, 
s.siteid + ' - ' + s.sitename as sitename,
isnull(c.amount,0) amount,
isnull(case when h.householdid is null then c.amount else h.amount end,0) as householdamount

from
anchor a 
left join constituent_amt c on a.constituentdimid = c.constituentdimid and a.siteid = c.siteid
left join household_amt h on a.householdid = h.householdid and a.siteid = h.siteid
left join bbinfinity_rpt_bbdw.bbdw.dim_site s on a.siteid = s.sitesystemid

where
a.householdid is not null

union

select
c.constituentdimid,
s.sitedimid, 
s.siteid + ' - ' + s.sitename as sitename,
isnull(c.amount, 0),
isnull(c.amount,0) householdamount

from
constituent_amt c
left join bbinfinity_rpt_bbdw.bbdw.dim_site s on c.siteid = s.sitesystemid

where
c.householdid is null