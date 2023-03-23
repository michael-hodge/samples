	
      

-- get million dollar philanthropist
with million as
(
select distinct cn.constituentid
from constituentnote cn join constituentnotetypecode cntc on cn.constituentnotetypecodeid = cntc.id
where cntc.description = 'principal gift note' and cn.title = 'million dollar philanthropist'
),

-- get MED-SOM preferred year
med_preferred_year0 as
(
select distinct parentid as constituentid, value
from V_QUERY_ATTRIBUTED7F1E1863A834216A8C7E7174249B5C4
where value like 'class of%'
),

med_preferred_year as
(
select distinct
x1.constituentid,
stuff
((
select distinct cast(',' as varchar(max)) + rtrim(ltrim(x2.value))
from med_preferred_year0 x2
where x1.constituentid = x2.constituentid
for xml path('')
), 1, 1, '') as med_som_preferred_year

from 
med_preferred_year0 x1
),


-- get latest child unc grad year
childclass as
(
select 
c.constituentdimid, max(e.classof) as maxclass

from
bbinfinity_rpt_bbdw.bbdw.dim_constituent c
join bbinfinity_rpt_bbdw.bbdw.fact_constituentrelationship r on c.constituentdimid = r.constituentdimid
join bbinfinity_rpt_bbdw.bbdw.dim_education e on r.reciprocalconstituentdimid = e.constituentdimid
left join bbinfinity_rpt_bbdw.bbdw.dim_constituent c2 on r.reciprocalconstituentdimid = c.constituentdimid
left join bbinfinity_rpt_bbdw.bbdw.dim_constituentrelationshiptype t on r.reciprocalconstituentrelationshiptypedimid = t.constituentrelationshiptypedimid

where
e.educationinstitutionisaffliated = 1
and t.constituentrelationshiptype in ('child', 'daughter', 'son', 'Step-child', 'Step-daughter', 'Step-son')
and e.educationconstituencystatus = 'Currently attending'
and e.educationprogram = 'Bachelors degree'

group by
c.constituentdimid
),

-- get interactions where constituent is primary or participant
interactions as
(
select 
fi.constituentsystemid, fi.interactiondate, di.interactiontype, c.fullname as fundraisername

from 
bbinfinity_rpt_bbdw.bbdw.fact_interaction fi
join bbinfinity_rpt_bbdw.bbdw.dim_interaction di on fi.interactiondimid = di.interactiondimid
left join bbinfinity_rpt_bbdw.bbdw.dim_constituent c on fi.fundraisersystemid = c.constituentsystemid

where 
di.interactionstatus = 'Completed' 

union
 
select 
ip.participantconstituentsystemid, ip.interactiondate, di.interactiontype, c.fullname as fundraisername
 
from 
bbinfinity_rpt_bbdw.bbdw.fact_interactionparticipant ip
join bbinfinity_rpt_bbdw.bbdw.dim_interaction di on ip.interactiondimid = di.interactiondimid
left join bbinfinity_rpt_bbdw.bbdw.dim_constituent c on ip.fundraisersystemid = c.constituentsystemid

where 
di.interactionstatus = 'Completed' 
),

-- get date of last interactions by type
last_interactions as
(
select 
constituentsystemid,
max(case when interactiontype in ('Phone', 'Call') then interactiondate else null end) as last_call,
max(case when interactiontype like ('Personal Visit%') then interactiondate else null end) as last_visit,
max(case when interactiontype in ('Mail') then interactiondate else null end) as last_mail,
max(case when interactiontype in ('Electronic') then interactiondate else null end) as last_electronic,
max(case when interactiontype in ('Event') then interactiondate else null end) as last_event,
max(interactiondate) as last_interaction

from 
interactions

group by 
constituentsystemid
),

-- sequence interactions by date
last_interaction_seq as
(
select constituentsystemid, interactiondate, interactiontype, fundraisername, row_number() over(partition by constituentsystemid order by interactiondate desc) as seq
from interactions
),

-- sequence last visit by date
last_visit_seq as
(
select constituentsystemid, interactiondate, interactiontype, fundraisername, row_number() over(partition by constituentsystemid order by interactiondate desc) as seq
from interactions
where interactiontype like ('Personal Visit%') 
),

-- get first year of giving 
minfy as
(
select constituentdimid, min(fiscalyear) as minfy
from bbinfinity_rpt_bbdw.bbdw.fact_usr_unc_aggregategiving_by_fy
where householdamount > 0
group by constituentdimid
),

-- get last year of giving 
maxfy as
(
select constituentdimid, max(fiscalyear) as maxfy
from bbinfinity_rpt_bbdw.bbdw.fact_usr_unc_aggregategiving_by_fy
where householdamount > 0
group by constituentdimid
),

-- get constituents with active plans based on pmra logic
activeplan as
(
select distinct
p.constituentdimid, p.constituentsystemid

from
bbinfinity_rpt_bbdw.bbdw.dim_prospect p 
left join bbinfinity_rpt_bbdw.bbdw.dim_prospectplan pp on p.constituentdimid = pp.constituentdimid
left join bbinfinity_rpt_bbdw.bbdw.dim_fundraisersecondary s on s.prospectplandimid = pp.prospectplandimid and s.todate is not null

where
(pp.prospectplanisactive = 1 and (pp.primaryfundraiserdimid is not null or pp.secondaryfundraiserdimid is not null or s.prospectplandimid is not null))
or
(p.prospectmanagerenddate is not null)
),

inactive_reason as
(
select distinct 
c.id as constituentid, ir.description as inactive_reason

from 
bbinfinity.dbo.constituent c 
join bbinfinity.dbo.constituentinactivedetail id on c.id = id.id
join bbinfinity.dbo.constituentinactivityreasoncode ir on id.constituentinactivityreasoncodeid = ir.id
),

-- get commitment plans and pivot to columns
commitmentplans as
(
select 
pp.constituentsystemid, pp.prospectplanname

from
bbinfinity_rpt_bbdw.bbdw.dim_prospectplan pp

where
pp.prospectplantype = 'Commitment'
),

commitmentplans_seq as
(
select
constituentsystemid, prospectplanname, row_number() over(partition by constituentsystemid order by prospectplanname) as seq

from
commitmentplans
),

commitmentplans_pivot as
(
select
constituentsystemid, 
max(case when seq = 1 then prospectplanname end) as plan1,
max(case when seq = 2 then prospectplanname end) as plan2 

from 
commitmentplans_seq

group by
constituentsystemid
),

-- get top 3 site giving and pivot to columns
sitegiving as
(
select constituentdimid, sitename, sum(householdamount) hhamt
from bbinfinity_rpt_bbdw.bbdw.fact_usr_unc_aggregategiving_by_site
group by constituentdimid, sitename
),

sitegiving_seq as
(
select
constituentdimid, sitename, hhamt, row_number() over(partition by constituentdimid order by hhamt desc) as seq

from
sitegiving
),

sitegiving_pivot as
(
select
constituentdimid,
max(case when seq = 1 then sitename end) as site1,
max(case when seq = 2 then sitename end) as site2, 
max(case when seq = 3 then sitename end) as site3, 
max(case when seq = 1 then hhamt end) as siteamt1,
max(case when seq = 2 then hhamt end) as siteamt2, 
max(case when seq = 3 then hhamt end) as siteamt3

from 
sitegiving_seq

group by
constituentdimid
),

-- get active plan sites
activeplan_sites as
(
select distinct
pp.constituentdimid, s.sitename

from
bbinfinity_rpt_bbdw.bbdw.dim_prospectplan pp 
join bbinfinity_rpt_bbdw.bbdw.fact_prospectplansite ps on pp.prospectplandimid = ps.prospectplandimid
join bbinfinity_rpt_bbdw.bbdw.dim_site s on ps.sitedimid = s.sitedimid

where
pp.prospectplanisactive = 1
),

activeplan_sites_seq as
(
select
constituentdimid, sitename, row_number() over(partition by constituentdimid order by sitename) as seq

from
activeplan_sites
),

activeplan_sites_pivot as
(
select
constituentdimid,
max(case when seq = 1 then sitename end) as activeplan_site1,
max(case when seq = 2 then sitename end) as activeplan_site2, 
max(case when seq = 3 then sitename end) as activeplan_site3,
max(case when seq = 4 then sitename end) as activeplan_site4,
max(case when seq = 5 then sitename end) as activeplan_site5,
max(case when seq = 6 then sitename end) as activeplan_site6,
max(case when seq = 7 then sitename end) as activeplan_site7,
max(case when seq = 8 then sitename end) as activeplan_site8

from 
activeplan_sites_seq

group by
constituentdimid
),

--get daf info
daf as
(
select distinct
f.constituentdimid, d.constituency

from 
bbinfinity_rpt_bbdw.bbdw.fact_constituency f
join bbinfinity_rpt_bbdw.bbdw.dim_constituency d on f.constituencydimid = d.constituencydimid

where
f.iscurrentconstituency = 1
and d.constituency in ('donor advised fund', 'community foundation', 'family foundation')
),

-- get total hh recognition years
hh_recog_years as
(
select constituentdimid, count(*) cnt
from bbinfinity_rpt_bbdw.bbdw.fact_usr_unc_aggregategiving_by_fy
where householdamount > 0
group by constituentdimid
)

select
d.CONSTITUENTSYSTEMID,
d.CONSTITUENTLOOKUPID,
d.HOUSEHOLDLOOKUPID,
case 
when isnull(d.HOUSEHOLDLOOKUPID, '') = '' then d.CONSTITUENTLOOKUPID
else d.HOUSEHOLDLOOKUPID
end as HOUSEHOLDDEDUPEID,
d.CONSTITUENTFULLNAME,
d.CONSTITUENTLASTNAME,
c.ISDECEASED,
ir.INACTIVE_REASON,
d.PRIMARYCONSTITUENCY,
c.AGE,
d.PRIMARYCLASS,
d.GENDER,
d.SPOUSECONSTITUENTLOOKUPID,
d.SPOUSEFULLNAME,
d.SPOUSEPRIMARYCONSTITUENCY,
d.PRIMARYADDRESSLINE1,
d.PRIMARYADDRESSLINE2,
d.PRIMARYADDRESSLINE3,
d.PRIMARYADDRESSCITY,
d.PRIMARYADDRESSSTATE,
d.PRIMARYADDRESSCOUNTY,
d.PRIMARYADDRESSPOSTCODE,
d.PRIMARYADDRESSCOUNTRY,
ac.latitude as PRIMARYADDRESSLATITUDE,
ac.longitude as PRIMARYADDRESSLONGITUDE,
agg_alltime.householdamount as AGGREGATE_ALLTIME_HOUSEHOLD,
agg_pg.householdirrevocableamount as AGGREGATE_PG_IRREVOCABLE_HOUSEHOLD,
agg_pg.householdamount as AGGREGATE_PG_RREVOCABLE_HOUSEHOLD,
agg_cmpn.householdamount as AGGREGATE_CAMPAIGN_HOUSEHOLD,
minfy.minfy as FIRST_FY_GIVING_HOUSEHOLD,
maxfy.maxfy as LAST_FY_GIVING_HOUSEHOLD,
agg_alltime.TOTAL_RECOG_YEARS,
hhry.cnt as HH_TOTAL_RECOG_YEARS,
isnull(d.PROSPECTMANAGER, '') as PROSPECTMANAGER,
isnull(d2.prospectmanager, '') as SPOUSE_PROSPECTMANAGER,
case when ap1.constituentsystemid is null then 'No' else 'Yes' end as ACTIVEPLAN,
case when ap1.constituentsystemid is null and ap2.constituentsystemid is null then 'No' else 'Yes' end as ACTIVEPLAN_HOUSEHOLD, 
--Plan Manager, HH,
cc.maxclass as LATEST_CHILD_CLASSYEAR,
cp1.plan1 as COMMITMENTPLAN1,
cp1.plan2 as COMMITMENTPLAN2,
cp2.plan1 as SPOUSE_COMMITMENTPLAN1,
cp2.plan2 as SPOUSE_COMMITMENTPLAN2,
lis1.interactiontype as LAST_INTERACTION_TYPE,
lis1.interactiondate as LAST_INTERACTION_DATE,
li1.last_call as LAST_CALL_DATE,
li1.last_visit as LAST_VISIT_DATE,
li1.last_mail as LAST_MAIL_DATE,
li1.last_electronic as LAST_ELECTRONIC_DATE,
li1.last_event as LAST_EVENT_DATE,
lis2.interactiontype as LAST_INTERACTION_TYPE_SPOUSE,
lis2.interactiondate as LAST_INTERACTION_DATE_SPOUSE,
li2.last_call as LAST_CALL_DATE_SPOUSE,
li2.last_visit as LAST_VISIT_DATE_SPOUSE,
li2.last_mail as LAST_MAIL_DATE_SPOUSE,
li2.last_electronic as LAST_ELECTRONIC_DATE_SPOUSE,
li2.last_event as LAST_EVENT_DATE_SPOUSE,
lis1.fundraisername as LAST_INTERACTION_OWNER,
lis2.fundraisername as LAST_INTERACTION_OWNER_SPOUSE,
liv1.fundraisername as LAST_VISIT_OWNER,
liv2.fundraisername as LAST_VISIT_OWNER_SPOUSE,
case when isnull(lis1.interactiondate, 0) >= isnull(lis2.interactiondate, 0) then lis1.interactiontype else lis2.interactiontype end as LAST_INTERACTION_TYPE_HOUSEHOLD,
case when isnull(lis1.interactiondate, 0) >= isnull(lis2.interactiondate, 0) then lis1.interactiondate else lis2.interactiondate end as LAST_INTERACTION_DATE_HOUSEHOLD,
case when isnull(li1.last_visit, 0) >= isnull(li2.last_visit, 0) then li1.last_visit else li2.last_visit end as LAST_VISIT_DATE_HOUSEHOLD,
case when isnull(li1.last_call, 0) >= isnull(li2.last_call, 0) then li1.last_call else li2.last_call end as LAST_CALL_DATE_HOUSEHOLD,
case when isnull(li1.last_electronic, 0) >= isnull(li2.last_electronic, 0) then li1.last_electronic else li2.last_electronic end as LAST_ELECTRONIC_DATE_HOUSEHOLD,
case when isnull(li1.last_mail, 0) >= isnull(li2.last_mail, 0) then li1.last_mail else li2.last_mail end as LAST_MAIL_DATE_HOUSEHOLD,
case when isnull(li1.last_event, 0) >= isnull(li2.last_event, 0) then li1.last_event else li2.last_event end as LAST_EVENT_DATE_HOUSEHOLD,
case when isnull(lis1.interactiondate, 0) >= isnull(lis2.interactiondate, 0) then lis1.fundraisername else lis2.fundraisername end as LAST_INTERACTION_OWNER_HOUSEHOLD,
case when isnull(liv1.interactiondate, 0) >= isnull(liv2.interactiondate, 0) then liv1.fundraisername else liv2.fundraisername end as LAST_VISIT_OWNER_HOUSEHOLD,
ds.ExactNearGiftCapacityRating as DONORSCAPE_RATING,
d.prospectstatus as UNC_PROSPECT_STATUS,
rfm.rfm as UNC_RFM,
z.unc_zip_wealth_score as UNC_ZIP_WEALTH_SCORE,
acot.total_score as UNC_AFFINITY_SCORE,
ss.engagement as SIMPSON_SCARBOROUGH_ENGAGEMENT,
ss.segment as SIMPSON_SCARBOROUGH_SEGMENT,
pgp.specialcode as PGP,
ewp.specialcode as EWP,
sg.site1 as TOP3_SITE1,
sg.site2 as TOP3_SITE2,
sg.site3 as TOP3_SITE3,
sg.siteamt1 as TOP3_SITEAMT1,
sg.siteamt2 as TOP3_SITEAMT2,
sg.siteamt3 as TOP3_SITEAMT3,
ap.ACTIVEPLAN_SITE1,
ap.ACTIVEPLAN_SITE2,
ap.ACTIVEPLAN_SITE3,
ap.ACTIVEPLAN_SITE4,
ap.ACTIVEPLAN_SITE5,
ap.ACTIVEPLAN_SITE6,
ap.ACTIVEPLAN_SITE7,
ap.ACTIVEPLAN_SITE8,
d.FOOTBALL as FOOTBALL_SEASON,
d.MENBASKETBALL as MENSBASKETBALL_SEASON,
d.WBASKETBALL as WOMENBASKETBALL_SEASON,
d.BASEBALL as BASEBALL_SEASON,
mpy.med_som_preferred_year as MED_SOM_PREFERRED_YEAR,
case when daf.constituentdimid is null then 0 else 1 end as DONOR_ADVISED_FUND,
case when cf.constituentdimid is null then 0 else 1 end as COMMUNITY_FOUNDATION,
case when ff.constituentdimid is null then 0 else 1 end as FAMILY_FOUNDATION,
case when m.constituentid is null then 0 else 1 end as MILLION_DOLLAR_PHILANTHROPIST,
'https://davie-crm.dev.unc.edu/bbappfx/webui/webshellpage.aspx?databasename=BBInfinity#pageType=p&pageId=88159265-2b7e-4c7b-82a2-119d01ecd40f&recordId=' +  cast(d.constituentsystemid as varchar(36)) as RECORD_URL

from
bbinfinity_rpt_bbdw.bbdw.fact_usr_unc_demographic d
join bbinfinity_rpt_bbdw.bbdw.dim_constituent c on d.constituentsystemid = c.constituentsystemid
left join bbinfinity.dbo.address a on d.constituentsystemid = a.constituentid and a.isprimary = 1
left join bbinfinity.dbo.addresscoordinates ac on a.id = ac.addressid
left join bbinfinity_rpt_bbdw.bbdw.fact_usr_unc_aggregategiving_alltime agg_alltime on d.constituentdimid = agg_alltime.constituentdimid
left join bbinfinity_rpt_bbdw.bbdw.fact_usr_unc_aggregateplannedgiving_alltime agg_pg on d.constituentdimid = agg_pg.constituentdimid
left join bbinfinity_rpt_bbdw.bbdw.fact_usr_unc_aggregatecampaigngiving_alltime agg_cmpn on d.constituentdimid = agg_cmpn.constituentdimid
left join minfy on minfy.constituentdimid = d.constituentdimid
left join maxfy on maxfy.constituentdimid = d.constituentdimid
left join bbinfinity_rpt_bbdw.bbdw.fact_usr_unc_demographic d2 on d.spouseconstituentsystemid = d2.constituentsystemid
left join activeplan ap1 on ap1.constituentsystemid = d.constituentsystemid
left join activeplan ap2 on ap2.constituentsystemid = d.spouseconstituentsystemid
left join bbinfinity_rpt_bbdw.dbo.usr_unc_pmra_rfmshortfile rfm on d.constituentsystemid = rfm.constituentid
left join bbinfinity_rpt_bbdw.bbdw.fact_usr_unc_acot acot on d.constituentsystemid = acot.constituentsystemid
left join last_interactions li1 on d.constituentsystemid = li1.constituentsystemid
left join last_interactions li2 on d.spouseconstituentsystemid = li2.constituentsystemid
left join last_interaction_seq lis1 on d.constituentsystemid = lis1.constituentsystemid and lis1.seq = 1
left join last_interaction_seq lis2 on d.spouseconstituentsystemid = lis2.constituentsystemid and lis2.seq = 1
left join last_visit_seq liv1 on d.constituentsystemid = liv1.constituentsystemid and liv1.seq = 1
left join last_visit_seq liv2 on d.spouseconstituentsystemid = liv2.constituentsystemid and liv2.seq = 1
left join bbinfinity_rpt_bbdw.dbo.dim_usr_unc_zip_score z on left(d.primaryaddresspostcode, 5) = left(z.zip_code, 5)
left join bbinfinity.dbo.donorscape2019 ds on ds.constituentid = d.constituentsystemid
left join inactive_reason ir on d.constituentsystemid = ir.constituentid
left join commitmentplans_pivot cp1 on d.constituentsystemid = cp1.constituentsystemid
left join commitmentplans_pivot cp2 on d.spouseconstituentsystemid = cp2.constituentsystemid
left join sitegiving_pivot sg on d.constituentdimid = sg.constituentdimid
left join activeplan_sites_pivot ap on d.constituentdimid = ap.constituentdimid
left join childclass cc on d.constituentdimid = cc.constituentdimid
left join med_preferred_year mpy on d.constituentsystemid = mpy.constituentid
left join daf on d.constituentdimid = daf.constituentdimid and daf.constituency = 'donor advised fund'
left join daf cf on d.constituentdimid = cf.constituentdimid and cf.constituency = 'community foundation' 
left join daf ff on d.constituentdimid = ff.constituentdimid and ff.constituency = 'family foundation'
left join hh_recog_years hhry on hhry.constituentdimid = d.constituentdimid
left join million m on d.constituentsystemid = m.constituentid

outer apply
(
select top 1 segment, engagement
from bbinfinity_rpt_bbdw.bbdw.dim_usr_unc_simpsonscarborough_segments ss
where d.constituentsystemid = ss.constituentsystemid
) ss

outer apply
(
select top 1 
fsc.constituentsystemid, dsc.specialcode 

from 
bbinfinity_rpt_bbdw.bbdw.fact_usr_unc_constituentspecialcode fsc join bbinfinity_rpt_bbdw.bbdw.dim_usr_unc_specialcode dsc on fsc.specialcodedimid = dsc.specialcodedimid

where
dsc.specialcode  in ('PGP', 'FPG', '1PG') and d.constituentsystemid = fsc.constituentsystemid
) pgp

outer apply
(
select top 1 
fsc.constituentsystemid, dsc.specialcode 

from 
bbinfinity_rpt_bbdw.bbdw.fact_usr_unc_constituentspecialcode fsc join bbinfinity_rpt_bbdw.bbdw.dim_usr_unc_specialcode dsc on fsc.specialcodedimid = dsc.specialcodedimid

where
dsc.specialcode  in ('EWP') and d.constituentsystemid = fsc.constituentsystemid
) ewp

		

