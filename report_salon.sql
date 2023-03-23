
create procedure [dbo].[USR_UNC_USP_SALON_EVENT]

(
@vGetImport as nvarchar(3),
@vCity as nvarchar(max),
@vState as nvarchar(2),
@vCounty as nvarchar(max),
@vRegion as nvarchar(max)
)

as begin

-- get current fiscal year
declare @CurrentFY smallint;
set @CurrentFY = (select 2000 + y.yearid fiscalyear from dbo.glfiscalperiod p join glfiscalyear y on p.glfiscalyearid = y.id where getdate() between p.startdate and p.enddate);


-- split city and county selections into tables
if object_id('tempdb.dbo.#salon_county') is not null
drop table #salon_county; 

create table #salon_county (val nvarchar(100))
insert into #salon_county (val) 
select val from bbinfinity.dbo.USR_UNC_UFN_STRING_TO_TABLE (@vCounty,',',1);

if object_id('tempdb.dbo.#salon_city') is not null
drop table #salon_city; 


create table #salon_city (val nvarchar(100))
insert into #salon_city (val) 
select val from bbinfinity.dbo.USR_UNC_UFN_STRING_TO_TABLE (@vCity,',',1);


if object_id('tempdb.dbo.#salon_region') is not null
drop table #salon_region; 


create table #salon_region (val nvarchar(100))
insert into #salon_region (val) 
select val from bbinfinity.dbo.USR_UNC_UFN_STRING_TO_TABLE (@vRegion,',',1);



-- get import records
declare @vImportName as varchar(50);
select @vImportName = case when @vGetImport = 'Yes' then 'UDO FYXX Salon Event' else null end;

declare @vImportID as uniqueidentifier;
set @vImportID = (select id from idsetregister where name like @vImportName + '%');

select 
c.constituentsystemid, c.constituentdimid, c.spouseconstituentdimid

into 
#ids_import

from 
dbo.ufn_idsetreader_getresults_guid(@vImportID) i
join bbinfinity_rpt_bbdw.bbdw.dim_constituent c on i.id = c.constituentsystemid;


-- get location records
select distinct
p.constituentsystemid, p.constituentdimid, c.spouseconstituentdimid

into #ids_location

from
bbinfinity_rpt_bbdw.bbdw.dim_prospect p
join bbinfinity_rpt_bbdw.bbdw.fact_constituentaddress fa on fa.constituentdimid = p.constituentdimid
join bbinfinity_rpt_bbdw.bbdw.dim_constituentaddress da on fa.constituentaddressdimid = da.constituentaddressdimid
join bbinfinity_rpt_bbdw.bbdw.dim_constituentaddressdetail d on fa.constituentaddressdetaildimid = d.constituentaddressdetaildimid
join bbinfinity_rpt_bbdw.bbdw.dim_constituentaddresstype t on fa.constituentaddresstypedimid = t.constituentaddresstypedimid
join bbinfinity_rpt_bbdw.bbdw.dim_constituentaddressflag f on fa.constituentaddressflagdimid = f.constituentaddressflagdimid
join bbinfinity_rpt_bbdw.bbdw.dim_constituent c on p.constituentdimid = c.constituentdimid
left join bbinfinity_rpt_bbdw.bbdw.dim_usr_unc_simpsonscarborough_segments ss on p.constituentdimid = ss.constituentdimid 
left join bbinfinity_rpt_bbdw.bbdw.fact_usr_unc_demographic_ext dext on c.constituentdimid = dext.constituentdimid 
left join bbinfinity.dbo.usr_unc_constituent_ext_gp gp on p.constituentsystemid = gp.id 
left join #salon_county county on county.val = d.county
left join #salon_city city on city.val = da.city
left join 
(
select pp.constituentdimid
from bbinfinity_rpt_bbdw.bbdw.dim_prospectplan pp
join bbinfinity_rpt_bbdw.bbdw.fact_opportunity o on pp.prospectplandimid = o.prospectplandimid
group by pp.constituentdimid
having sum(opportunityamount) >= 100000
) opp on p.constituentdimid = opp.constituentdimid


where
da.stateabbreviation = @vState 
and ((county.val is not null or city.val is not null) or (@vCounty = '' and @vCity = ''))
and isnull(gp.patientonly,0) = 0
and c.isdeceased = 0
and c.isactive = 1
and c.isindividual = 1
and (f.isprimary = 1 or t.constituentaddresstype in ('primary', 'seasonal','alternate','other'))
and fa.historicalenddate is null
and da.countryabbreviation = 'usa'
and 
(
opp.constituentdimid is not null
or
ss.segment like '%rockefellers%'
or 
ss.segment like '%gatsbys%'
or 
ss.segment like '%undercover bosses%'
or 
ss.segment like '%tar heels for life%'
or
(p.prospectstatus Not like ('x%') and p.prospectstatus <> 'No prospect status')
or
(dext.primarygaamembershipprogram = 'gaa: lifetime membership' and dext.primarygaamembershipexpirationdate is null)
)
;


-- get region records
select distinct
p.constituentsystemid, p.constituentdimid, c.spouseconstituentdimid

into #ids_region

from
bbinfinity_rpt_bbdw.bbdw.dim_prospect p
join bbinfinity_rpt_bbdw.bbdw.fact_constituentaddress fa on fa.constituentdimid = p.constituentdimid
join bbinfinity_rpt_bbdw.bbdw.dim_constituentaddress da on fa.constituentaddressdimid = da.constituentaddressdimid
join bbinfinity_rpt_bbdw.bbdw.dim_constituentaddressdetail d on fa.constituentaddressdetaildimid = d.constituentaddressdetaildimid
join bbinfinity_rpt_bbdw.bbdw.dim_constituentaddresstype t on fa.constituentaddresstypedimid = t.constituentaddresstypedimid
join bbinfinity_rpt_bbdw.bbdw.dim_constituentaddressflag f on fa.constituentaddressflagdimid = f.constituentaddressflagdimid
join bbinfinity_rpt_bbdw.bbdw.dim_constituent c on p.constituentdimid = c.constituentdimid
left join bbinfinity_rpt_bbdw.bbdw.dim_usr_unc_simpsonscarborough_segments ss on p.constituentdimid = ss.constituentdimid 
left join bbinfinity_rpt_bbdw.bbdw.fact_usr_unc_demographic_ext dext on c.constituentdimid = dext.constituentdimid
left join bbinfinity.dbo.usr_unc_constituent_ext_gp gp on p.constituentsystemid = gp.id 
join bbinfinity.dbo.usr_constituentcustomcodes cc on fa.constituentsystemid = cc.constituentid  and fa.constituentaddresssystemid = cc.addressid
join bbinfinity.dbo.usr_customcodetypecodevalue cv on cc.customcodetypevalueid = cv.id
left join #salon_region region on region.val = cv.description

where
region.val is not null
and isnull(gp.patientonly,0)=0
and c.isdeceased = 0
and c.isactive = 1
and (f.isprimary = 1 or t.constituentaddresstype in ('primary', 'seasonal','alternate','other'))
and fa.historicalenddate is null
and da.countryabbreviation = 'usa'
and 
(
ss.segment like '%rockefellers%'
or 
ss.segment like '%gatsbys%'
or 
ss.segment like '%undercover bosses%'
or 
ss.segment like '%tar heels for life%'
or
(p.prospectstatus Not like ('x%') and p.prospectstatus <> 'No prospect status')
or
(dext.primarygaamembershipprogram = 'gaa: lifetime membership' and dext.primarygaamembershipexpirationdate is null)
)


-- merge ids from location, region, and import selections
select distinct *
into #ids
from (select * from #ids_import union select * from #ids_location union select * from #ids_region) x


create nonclustered index ix_tempid1 on #ids (constituentdimid);
create nonclustered index ix_tempid2 on #ids (constituentsystemid);
create nonclustered index ix_tempid3 on #ids (spouseconstituentdimid);



-- get interactions where constituent is primary or participant
select * into #interactions from
(
select 
fi.constituentdimid, fi.interactiondate, di.interactiontype, c.fullname as fundraisername,
case when di.interactiontype like 'personal%' then 1 else 0 end as visit_ind

from 
bbinfinity_rpt_bbdw.bbdw.fact_interaction fi
join bbinfinity_rpt_bbdw.bbdw.dim_interaction di on fi.interactiondimid = di.interactiondimid
join bbinfinity_rpt_bbdw.bbdw.dim_constituent c on fi.fundraiserdimid = c.constituentdimid
join #ids on #ids.constituentdimid = fi.constituentdimid

where 
di.interactionstatus = 'completed' 

union
 
select 
ip.participantconstituentdimid, ip.interactiondate, di.interactiontype, c.fullname as fundraisername,
case when di.interactiontype like 'personal%' then 1 else 0 end as visit_ind
 
from 
bbinfinity_rpt_bbdw.bbdw.fact_interactionparticipant ip
join bbinfinity_rpt_bbdw.bbdw.dim_interaction di on ip.interactiondimid = di.interactiondimid
join bbinfinity_rpt_bbdw.bbdw.dim_constituent c on ip.fundraiserdimid = c.constituentdimid
join #ids on #ids.constituentdimid = ip.participantconstituentdimid
 
where 
di.interactionstatus = 'completed' 
) x


-- determine last overall contact and last visit
select *, 
row_number() over(partition by constituentdimid order by interactiondate desc) as seq,
case when visit_ind = 0 then 0 else row_number() over(partition by constituentdimid, visit_ind order by interactiondate desc) end as visitseq
into #interactions_seq
from #interactions;

-- get last interactions
select *
into #last_interactions
from #interactions_seq 
where (seq = 1 or visitseq = 1);



-- get top 3 site giving and pivot to columns
select 
a.constituentdimid, s.sitename, sum(a.householdamount) hhamt

into #sitegiving

from 
bbinfinity_rpt_bbdw.bbdw.fact_usr_unc_aggregategiving_by_site a 
join bbinfinity_rpt_bbdw.bbdw.dim_site s on a.sitedimid = s.sitedimid
join #ids on #ids.constituentdimid = a.constituentdimid

group by 
a.constituentdimid, s.sitename;


select
constituentdimid, sitename, hhamt, row_number() over(partition by constituentdimid order by hhamt desc) as seq

into #sitegiving_seq

from
#sitegiving;


select
constituentdimid,
max(case when seq = 1 then sitename end) as site1,
max(case when seq = 2 then sitename end) as site2, 
max(case when seq = 3 then sitename end) as site3, 
max(case when seq = 1 then hhamt end) as siteamt1,
max(case when seq = 2 then hhamt end) as siteamt2, 
max(case when seq = 3 then hhamt end) as siteamt3

into #sitegiving_pivot

from 
#sitegiving_seq

group by
constituentdimid;



-- get top 3 campaign site giving and pivot to columns
select 
a.constituentdimid, s.sitename, sum(a.householdamount) hhamt

into #camp_sitegiving

from 
bbinfinity_rpt_bbdw.bbdw.fact_usr_unc_aggregatecampaigngiving_by_site a 
join bbinfinity_rpt_bbdw.bbdw.dim_site s on a.sitedimid = s.sitedimid
join #ids on #ids.constituentdimid = a.constituentdimid

group by 
a.constituentdimid, s.sitename;


select
constituentdimid, sitename, hhamt, row_number() over(partition by constituentdimid order by hhamt desc) as seq

into #camp_sitegiving_seq

from
#camp_sitegiving;


select
constituentdimid,
max(case when seq = 1 then sitename end) as site1,
max(case when seq = 2 then sitename end) as site2, 
max(case when seq = 3 then sitename end) as site3, 
max(case when seq = 1 then hhamt end) as siteamt1,
max(case when seq = 2 then hhamt end) as siteamt2, 
max(case when seq = 3 then hhamt end) as siteamt3

into #camp_sitegiving_pivot

from 
#camp_sitegiving_seq

group by
constituentdimid;



-- get seaonal, alternate, and other addresses
select 
t.constituentaddresstype as addrtype, fa.constituentdimid, fa.addressblock, da.city, da.state, da.postcode, da.country, d.county, fa.historicalstartdate as addrstartdate, fa.historicalenddate as addrenddate

into #addr 

from 
bbinfinity_rpt_bbdw.bbdw.dim_constituentaddress da
join bbinfinity_rpt_bbdw.bbdw.fact_constituentaddress fa on fa.constituentaddressdimid = da.constituentaddressdimid
join bbinfinity_rpt_bbdw.bbdw.dim_constituentaddresstype t on fa.constituentaddresstypedimid = t.constituentaddresstypedimid
join bbinfinity_rpt_bbdw.bbdw.dim_constituentaddressdetail d on fa.constituentaddressdetaildimid = d.constituentaddressdetaildimid
join #ids on #ids.constituentdimid = fa.constituentdimid
left join #salon_county county on county.val = d.county
left join #salon_city city on city.val = da.city

where
da.stateabbreviation = @vState 
and ((county.val is not null or city.val is not null) or (@vCounty = '' and @vCity = ''))
and t.constituentaddresstype in ('seasonal','alternate','other' ) 
and fa.historicalenddate is null;


-- get campaign status
select 
p.id as prospectid, c.description as status

into #campaign_status

from 
bbinfinity.dbo.prospect p 
join bbinfinity.dbo.prospectstatuscode psc on p.prospectstatuscodeid = psc.id
join bbinfinity.dbo.usr_unc_prospect_ext pe on p.id = pe.id
join bbinfinity.dbo.usr_unc_campaignstatuscode c on pe.campaignstatuscodeid = c.id
join #ids on #ids.constituentsystemid = p.id;



-- get opportunity amount
select 
pp.constituentdimid, sum (opportunityamount) opportunityamount

into #prospectplan

from bbinfinity_rpt_bbdw.bbdw.dim_prospectplan pp
join bbinfinity_rpt_bbdw.bbdw.fact_opportunity o on pp.prospectplandimid = o.prospectplandimid
join #ids on #ids.constituentdimid = pp.constituentdimid

group by 
pp.constituentdimid

having 
sum(opportunityamount) >= 100000;



-- get planned gifts
select 
pg.constituentdimid 

into #plannedgift

from 
bbinfinity_rpt_bbdw.bbdw.dim_plannedgift pg join #ids on #ids.constituentdimid = pg.constituentdimid

where 
pg.status = 'accepted';



-- get last year of giving
select 
a.constituentdimid, max(a.fiscalyear) maxyear

into #last_year_giving

from 
bbinfinity_rpt_bbdw.bbdw.fact_usr_unc_aggregategiving_by_fy a
join #ids on #ids.constituentdimid = a.constituentdimid

where 
a.amount >  0

group by 
a.constituentdimid;


select 
a.constituentdimid, a.fiscalyear, a.amount

into #last_year_giving_amt

from 
bbinfinity_rpt_bbdw.bbdw.fact_usr_unc_aggregategiving_by_fy a 
join #last_year_giving l on a.constituentdimid = l.constituentdimid and a.fiscalyear = l.maxyear
join #ids on #ids.constituentdimid = a.constituentdimid

where 
a.amount > 0;



-- get constituencies
select distinct 
f.constituentdimid, d.constituency

into #constituency

from bbinfinity_rpt_bbdw.bbdw.fact_constituency f 
join bbinfinity_rpt_bbdw.bbdw.dim_constituency d on f.constituencydimid = d.constituencydimid
join #ids on #ids.constituentdimid = f.constituentdimid

where 
f.todate is null;


-- get primary education
select
e.constituentdimid, 
e.constituentsystemid, 
e.educationdegree as primarydegree, 
e.classof as primaryclass,
min(ei.educationalcollege) as primarycollege, 
min(ei.educationaldepartment) as primarydept

into #education

from 
bbinfinity_rpt_bbdw.bbdw.dim_education e 
join bbinfinity_rpt_bbdw.bbdw.dim_educationadditionalinformation ei on e.educationdimid = ei.educationdimid
left join #ids ids1 on e.constituentdimid = ids1.constituentdimid
left join #ids ids2 on e.constituentdimid = ids2.spouseconstituentdimid

where 
e.isprimaryrecord = 1
and (ids1.constituentdimid is not null or ids2.spouseconstituentdimid is not null)

group by
e.constituentdimid, 
e.constituentsystemid, 
e.educationdegree, 
e.classof;


-- get board members
select distinct 
c2.constituentdimid, max(isnull(m.todate,'9999-01-01')) as enddate,
c.constituentlookupid, cg.grouptype, c.keyname

into #board

from 
bbinfinity_rpt_bbdw.bbdw.dim_constituentgroup cg
join bbinfinity_rpt_bbdw.bbdw.dim_constituent c on cg.constituentdimid = c.constituentdimid
join bbinfinity_rpt_bbdw.bbdw.fact_constituentgroupmember m on cg.constituentgroupdimid = m.constituentgroupdimid
join bbinfinity_rpt_bbdw.bbdw.dim_constituent c2 on m.memberconstituentdimid = c2.constituentdimid

where
cg.grouptype in ('board', 'committee')
and c.keyname in 
(
'unc board of trustees',
'chancellor’s philanthropic council',
'chancellor’s philanthropic council emeriti',
'campaign for carolina steering committee',
'carolina 1st campaign cabinet members 98-99'
)

group by
c2.constituentdimid, c.constituentlookupid, cg.grouptype, c.keyname;


-- results
select distinct
'https://davie-crm.dev.unc.edu/bbappfx/webui/webshellpage.aspx?databasename=BBInfinity#pageType=p&pageId=88159265-2b7e-4c7b-82a2-119d01ecd40f&recordId=' + convert(varchar(36), d.constituentsystemid)  as url,
d.constituentlookupid as 'unc_id',
d.householdlookupid as 'household_id',
case when d.householdlookupid <> '' then d.householdlookupid else d.constituentlookupid end as 'household_dedupe_id',
case when c.isdeceased = 1 then 'Yes' else 'No' end as 'deceased',
case when c.isactive = 0 then 'Yes' else 'No' end as 'inactive',
case when d.isorganization = 1 then 'No' else 'Yes' end as individual,
d.primaryconstituency as 'primary_constituency',
case when facstaff.constituentdimid is null then '' else 'Faculty Staff' end as 'faculty_staff',
case when formerfacstaff.constituentdimid is null then '' else 'Former Faculty Staff' end as 'former_faculty_staff',
case when student.constituentdimid is null then '' else 'Parent Current Student' end as 'parent_current_student',
case when formerstudent.constituentdimid is null then '' else 'Parent Former Student' end as 'parent_former_student',
case when friend.constituentdimid is null then '' else 'Friend' end as 'friend',
c.title as 'title',
d.constituentlastname as 'lastname',
d.constituentfirstname as 'firstsname',
d.constituentnickname as 'nickname',
d.constituentfullname as 'fullname',
case when d.isaNonymous = 1 then 'Yes' else 'No' end as 'givesanonymously',
d.spouseconstituentlookupid as 'spouse_id',
d.spousefullname as 'spousename',
case when c2.isdeceased = 1 then 'Yes' else 'No' end as 'spouse_deceased',
case when c2.isactive = 0 then 'Yes' else 'No' end as 'spouse_inactive',
d.primaryaddress as 'address',
d.primaryaddresscity as 'city',
d.primaryaddressstate as 'state',
d.primaryaddresscountry as 'country',
d.primaryaddresscounty as 'county',
d.primaryaddresspostcode as 'zip',
d.ethnicity as 'ethnicity',
d.constituentbirthdate as 'birthdate',
c.age as 'age',
d.gender as 'gender',
case when dext.constituentdimid is null then '' else 'Active' end as 'gaa_lifetime_status',
p.prospectstatus as 'unc_rating',
case 
when p.prospectstatus like 'aaaaa%' then 100000000
when p.prospectstatus like 'aaaa%' then 50000000
when p.prospectstatus like 'aaa%' then 25000000
when p.prospectstatus like 'aa%' then 10000000
when p.prospectstatus like 'a%' then 5000000
when p.prospectstatus like 'b%' then 1000000
when p.prospectstatus like 'c%' then 500000
when p.prospectstatus like 'd%' then 100000
when p.prospectstatus like 'e%' then 25000
when p.prospectstatus like 'xaaaaa%' then 100000000
when p.prospectstatus like 'xaaaa%' then 50000000
when p.prospectstatus like 'xaaa%' then 25000000
when p.prospectstatus like 'xaa%' then 10000000
when p.prospectstatus like 'xa%' then 5000000
when p.prospectstatus like 'xb%' then 1000000
when p.prospectstatus like 'xc%' then 500000
when p.prospectstatus like 'xd%' then 100000
when p.prospectstatus like 'xe%' then 25000
else 0 end as 'unc_capacity',
d.prospectmanager as 'prospect_manager',
cs.status as 'campaign_status',
a.total_score as 'unc_affinity_score',
ds.exactneargiftcapacityrating as 'donorscape',
(select top 1 c.description
from bbinfinity.dbo.attribute211d1d9e300f4e438330b676977ec212 r 
join bbinfinity.dbo.ggagiftcapacityratingcode c on c.id = r.ggagiftcapacityratingcodeid
where r.modelingandpropensityid = d.constituentsystemid) as 'gga_gift_capacity',
ss.segment as 'simpson_scarborough',
case when pp.constituentdimid is null then 'No' else 'Yes' end as 'active_plan',
case when isnull(pp.opportunityamount, 0) >  0 then 'Yes' else 'No' end as 'opportunity',
case when pp.constituentdimid is null then 'No' else 'Yes' end as 'planned_gift',
lc.interactiondate as 'last_contact_date',
lc.interactiontype as 'last_contact_method',
lv.interactiondate as 'last_visit_date',
lv.fundraisername as 'last_visit_fundraiser',
(select o.alltimeamount from bbinfinity.dbo.usr_unc_v_query_overallalltimegiving o where o.id = #ids.constituentsystemid) as 'overall_commitment',
agg.amount as 'alltime_recognition',
aggcamp.amount as 'campaign_recognition_giving',
agglegal.amount as 'camp_legal_giving', 
aggcamp.householdamount as 'campaign_household_recognition',
agg.householdamount as 'household_alltime_recognition',
(select o.alltimeamount from bbinfinity.dbo.usr_unc_v_query_overallalltimegiving o where o.id = d.spouseconstituentsystemid) as 'spouse_overall_commitment',
aggspouse.amount as 'spouse_alltime_recognition',
aggcampspouse.amount as 'spouse_campaign_recognition_giving',
agglegalspouse.amount as 'spouse_camp_legal_giving', 
aggcamp.householdamount as 'spouse_campaign_household_recognition',
agg.householdamount as 'spouse_household_alltime_recognition',
curfy.amount as 'curfy_amount',
ly.fiscalyear as 'last_fy_giving',
ly.amount as 'last_fy_amount',
case when b1.constituentdimid is null then '' else 'UNC Board of Trustees' end as 'unc_board_of_trustees',
b1.enddate as 'bot_term_end_date',
case when b2.constituentdimid is null then '' else 'Chancellors Philanthropic Council' end as 'chancellors_philanthropic_council',
case when b3.constituentdimid is null then '' else 'Chancellors Philanthropic Council Emeriti' end as 'chancellors_philanthropic_council_emeriti',
case when b4.constituentdimid is null then '' else 'Campaign for Carolina Steering Committee' end as 'campaign_for_carolina_steering_committee',
case when b5.constituentdimid is null then '' else ' Carolina First Campaign Steering Committee' end as 'carolina_first_campaign_steering_committee',
case when pgp.constituentsystemid is null then '' else 'PGP' end as 'pgp',
case when ogp.constituentsystemid is null then '' else '1PG' end as '1pg',
case when fpg.constituentsystemid is null then '' else 'FPG' end as 'fpg',
case when ewp.constituentsystemid is null then '' else 'EWP' end as 'EWP',
case when owp.constituentsystemid is null then '' else '1WP' end as '1WP',
addr_seasonal.addressblock as 'seasonal_address',
addr_seasonal.city as 'seasonal_city',
addr_seasonal.state as 'seasonal_state',
addr_seasonal.postcode as 'seasonal_zip',
addr_seasonal.county as 'seasonal_county',
addr_seasonal.country as 'seasonal_country',
addr_seasonal.addrstartdate as 'seasonal_start_date',
addr_seasonal.addrenddate as 'seasonal_end_date', 
addr_other.addressblock as 'other_address',
addr_other.city as 'other_city',
addr_other.state as 'other_state',
addr_other.postcode as 'other_zip',
addr_other.county as 'other_county',
addr_other.country as 'other_country',
addr_alternate.addressblock as 'alternate_address',
addr_alternate.city as 'alternate_city',
addr_alternate.state as 'alternate_state',
addr_alternate.postcode as 'alternate_zip',
addr_alternate.county as 'alternate_county',
addr_alternate.country as 'alternate_country',
rfm.rfm_rating as 'rfm_score',
d.employer as 'employer',
d.jobtitle as 'job_title',
z.unc_zip_wealth_score as 'zipwealth',
e.primarydegree as 'primary_degree',
e.primaryclass as 'primary_class',
e.primarycollege as 'primary_college',
e.primarydept as 'primary_department',
e2.primarydegree as 'spouse_primary_degree',
e2.primaryclass as 'spouse_primary_class',
e2.primarycollege as 'spouse_primary_college',
e2.primarydept as 'spouse_primary_department',
top3.site1 as 'top3_giving_sites1',
top3.siteamt1 as 'top3_giving_amount1',
top3.site2 as 'top3_giving_sites2',
top3.siteamt2 as 'top3_giving_amount2',
top3.site3 as 'top3_giving_sites3',
top3.siteamt3 as 'top3_giving_amount3',
top3spouse.site1 as 'spouse_top3_giving_sites1',
top3spouse.siteamt1 as 'spouse_top3_giving_amount1',
top3spouse.site2 as 'spouse_top3_giving_sites2',
top3spouse.siteamt2 as 'spouse_top3_giving_amount2',
top3spouse.site3 as 'spouse_top3_giving_sites3',
top3spouse.siteamt3 as 'spouse_top3_giving_amount3',
d.primaryemailaddress as 'primary_email',
d2.primaryemailaddress as 'spouse_primary_email',
d.uncjointformalsalutationname as 'joint_formal_salutation_name',
d.uncjointformaladdresseename as 'joint_formal_addressee_name',
d.primaryphonenumber as 'primary_phone',
top3camp.site1 as 'camp_top3_giving_sites1',
top3camp.siteamt1 as 'camp_top3_giving_amount1',
top3camp.site2 as 'camp_top3_giving_sites2',
top3camp.siteamt2 as 'camp_top3_giving_amount2',
top3camp.site3 as 'camp_top3_giving_sites3',
top3camp.siteamt3 as 'camp_top3_giving_amount3',
top3campspouse.site1 as 'spouse_camp_top3_giving_sites1',
top3campspouse.siteamt1 as 'spouse_camp_top3_giving_amount1',
top3campspouse.site2 as 'spouse_camp_top3_giving_sites2',
top3campspouse.siteamt2 as 'spouse_camp_top3_giving_amount2',
top3campspouse.site3 as 'spouse_camp_top3_giving_sites3',
top3campspouse.siteamt3 as 'spouse_camp_top3_giving_amount3',
isnull(a.total_score, 0) / 2 as affinity_score,

case 
when p.prospectstatus = 'a1 - $5,000,000-$9,999,999 confident' then 50.0
when p.prospectstatus = 'a2 - $5,000,000-$9,999,999 minimum threshold' then	50.0
when p.prospectstatus = 'a3 - $5,000,000-$9,999,999 unconfirmed/screening' then	50.0
when p.prospectstatus = 'aa1 - $10,000,000-$24,999,999 confident' then	50.0
when p.prospectstatus = 'aaa1 - $25,000,000-$49,999,999 confident' then	50.0
when p.prospectstatus = 'aaa3 - $25,000,000-$49,999,999 unconfirmed/screening' then	50.0
when p.prospectstatus = 'aaaa1 - $50,000,000-$99,999,999 confident' then	50.0
when p.prospectstatus = 'aaaa2 - $50,000,000-$99,999,999 minimum threshold' then 50.0
when p.prospectstatus = 'aaaa3 - $50,000,000-$99,999,999 unconfirmed/screening' then 50.0
when p.prospectstatus = 'aaaaa1 - $100,000,000+ confident' then	50.0
when p.prospectstatus = 'aaaa2 - $50,000,000-$99,999,999 reasonably confident' then	50.0
when p.prospectstatus = 'aaaa3 - $50,000,000-$99,999,999 unconfirmed' then	50.0
when p.prospectstatus = 'aaaaa1 - $100,000,000+ confirmed' then	50.0
when p.prospectstatus = 'aaaaa2 - $100,000,000+ reasonably confident' then	50.0
when p.prospectstatus = 'b1 - $1,000,000-$4,999,999 confident' then	47.5
when p.prospectstatus = 'b2 - $1,000,000-$4,999,999 minimum threshold' then	47.5
when p.prospectstatus = 'b3 - $1,000,000-$4,999,999 unconfirmed/screening' then	47.5
when p.prospectstatus = 'c1 - $500,000-$999,999 confident' then	45.0
when p.prospectstatus = 'c2 - $500,000-$999,999 minimum threshold' then	45.0
when p.prospectstatus = 'c3 - $500,000-$999,999 unconfirmed/screening' then	45.0
when p.prospectstatus = 'd0 - $100,000-$499,999 outdated' then	42.5
when p.prospectstatus = 'd1 - $100,000-$499,999 confident' then	42.5
when p.prospectstatus = 'd2 - $100,000-$499,999 minimum threshold' then	42.5
when p.prospectstatus = 'd3 - $100,000-$499,999 unconfirmed/screening' then	42.5
when p.prospectstatus = 'e1 - $25,000-$99,999 confident' then 25.0
when p.prospectstatus = 'e2 - $25,000-$99,999 minimum threshold' then 25.0
when p.prospectstatus = 'e3 - $25,000-$99,999 unconfirmed/screening' then 25.0
when p.prospectstatus = 'e0 - $25,000-$99,999 outdated' then 12.5
when ds.exactneargiftcapacityrating = '1 - more than $10 million' then 50.0
when ds.exactneargiftcapacityrating = '2 - $1 million to $9,999,999' then 47.5
when ds.exactneargiftcapacityrating = '3 - $250,000 to $999,999' then 45.0
when ds.exactneargiftcapacityrating = '4 - $100,000 to $249,999' then 42.5
when ds.exactneargiftcapacityrating = '5 - $25,000 to $99,999' then	25.0
when ds.exactneargiftcapacityrating = '6 - $10,000 to $24,999' then	10.0
when ds.exactneargiftcapacityrating = '7 - $2,500 to $9,999' then 7.5
when ds.exactneargiftcapacityrating = '8 - less than $2,500' then 5.0
else 0
end as capacity_score,

case 
when d.ethnicity like '%american indian or alaska native%' then 5.0
when d.ethnicity like '%asian or pacific islander%' then 5.0
when d.ethnicity like '%black or african american%' then 5.0
when d.ethnicity like '%hispanic or latiNo%' then 5.0
when d.ethnicity like '%asian%' then 5.0
else 0 end as bonus_ethnicity_score,

case 
when pgp.specialcode = 'pgp' then 5
when EWP.specialcode = 'EWP' then 5
when fpg.specialcode = 'fpg' then 5
else 0
end as bonus_specialcode_score,

case when d.prospectmanager is null then 0 else 5 end as bonus_mngr_score,

case
when ss.segment = 'rockefellers' then 5.0
when ss.segment = 'rockefellers - predicted' then 5.0
when ss.segment = 'gatsbys' then 4.5
when ss.segment = 'gatsbys - predicted' then 4.5
when ss.segment = 'tar heels for life' then	4.5
when ss.segment = 'tar heels for life - predicted' then	4.5
when ss.segment = 'undercover bosses' then 4.0
when ss.segment = 'undercover bosses - predicted' then 4.0
when ss.segment = 'untapped goldminers' then 4.0
when ss.segment = 'family focused' then	3.0
when ss.segment = 'family focused - predicted' then	3.0
when ss.segment = 'tailgaters' then	2.5
when ss.segment = 'tailgaters - predicted' then	2.5
when ss.segment = 'honeymooners or hipster parents' then 2.0
when ss.segment = 'honeymooners or hipster parents - predicted' then 2.0
when ss.segment = 'losing touchers or facebook moms' then 1.0
when ss.segment = 'losing touchers or facebook moms - predicted' then 1.0
when ss.segment = 'recent grads' then 1.0
when ss.segment = 'bar hoppers' then 0.5
when ss.segment = 'bar hoppers - predicted' then 0.5
when ss.segment = 'diverse young professionals' then 0.5
when ss.segment = 'diverse young professionals - predicted' then 0.5
when ss.segment = 'recent grads - predicted' then 0.5
else 0 end as bonus_ss_score,

case
when agg.amount >= 100000 then 10.00
when agg.amount >= 90000 then 7.50
when agg.amount >= 80000 then 6.75
when agg.amount >= 70000 then 6.00
when agg.amount >= 60000 then 5.25
when agg.amount >= 50000 then 4.50
when agg.amount >= 40000 then 3.75
when agg.amount >= 30000 then 3.00
when agg.amount >= 20000 then 2.25
when agg.amount >= 10000 then 1.50
when agg.amount >= 1000 then 0.75
when agg.amount > 0 then 0.10
else 0
end as bonus_giving_score

into 
#results

from
#ids
join bbinfinity_rpt_bbdw.bbdw.fact_usr_unc_demographic d on #ids.constituentdimid = d.constituentdimid
join bbinfinity_rpt_bbdw.bbdw.dim_constituent c on d.constituentdimid = c.constituentdimid
left join bbinfinity_rpt_bbdw.bbdw.dim_constituent c2 on c.spouseconstituentdimid = c2.constituentdimid
left join bbinfinity_rpt_bbdw.bbdw.fact_usr_unc_demographic d2 on c2.constituentdimid = d2.constituentdimid
left join bbinfinity_rpt_bbdw.bbdw.fact_usr_unc_demographic_ext dext on d.constituentdimid = dext.constituentdimid
join bbinfinity_rpt_bbdw.bbdw.dim_prospect p on d.constituentdimid = p.constituentdimid
left join #campaign_status cs on d.constituentsystemid = cs.prospectid
left join bbinfinity_rpt_bbdw.bbdw.fact_usr_unc_acot a on d.constituentdimid = a.constituentdimid
left join bbinfinity.dbo.doNorscape2019 ds on d.constituentsystemid = ds.constituentid
left join bbinfinity_rpt_bbdw.dbo.dim_usr_unc_zip_score z on z.zip_code = substring(d.primaryaddresspostcode,1,5)
left join bbinfinity_rpt_bbdw.bbdw.dim_usr_unc_simpsonscarborough_segments ss on d.constituentdimid = ss.constituentdimid
left join #prospectplan pp on d.constituentdimid = pp.constituentdimid
left join #plannedgift pg on d.constituentdimid = pg.constituentdimid
left join bbinfinity_rpt_bbdw.bbdw.fact_usr_unc_aggregatecampaigngiving_alltime aggcamp on d.constituentdimid = aggcamp.constituentdimid
left join bbinfinity_rpt_bbdw.bbdw.fact_usr_unc_aggregategiving_alltime agg on d.constituentdimid = agg.constituentdimid
left join bbinfinity_rpt_bbdw.bbdw.fact_usr_unc_aggregatecampaigngiving_alltime aggcampspouse on c.spouseconstituentdimid = aggcampspouse.constituentdimid
left join bbinfinity_rpt_bbdw.bbdw.fact_usr_unc_aggregategiving_alltime aggspouse on c.spouseconstituentdimid = aggspouse.constituentdimid
left join bbinfinity_rpt_bbdw.bbdw.fact_usr_unc_aggregategiving_by_fy curfy on d.constituentdimid = curfy.constituentdimid and fiscalyear = @CurrentFY
left join #last_year_giving_amt ly on d.constituentdimid = ly.constituentdimid
left join #constituency facstaff on facstaff.constituentdimid = d.constituentdimid and facstaff.constituency = 'faculty/staff'
left join #constituency formerfacstaff on formerfacstaff.constituentdimid = d.constituentdimid and formerfacstaff.constituency = 'former faculty/staff'
left join #constituency student on student.constituentdimid = d.constituentdimid and student.constituency = 'parent-current student'
left join #constituency formerstudent on formerstudent.constituentdimid = d.constituentdimid and formerstudent.constituency = 'parent-former student'
left join #constituency friend on friend.constituentdimid = d.constituentdimid and friend.constituency = 'friend'
left join #addr addr_seasonal on addr_seasonal.constituentdimid = d.constituentdimid and addr_seasonal.addrtype = 'seasonal'
left join #addr addr_alternate on addr_alternate.constituentdimid = d.constituentdimid and addr_alternate.addrtype = 'alternate'
left join #addr addr_other on addr_other.constituentdimid = d.constituentdimid and addr_other.addrtype = 'other'
left join bbinfinity_rpt_bbdw.dbo.usr_unc_pmra_rfmshortfile rfm on d.constituentlookupid = rfm.lookupid
left join #education e on #ids.constituentdimid = e.constituentdimid
left join #education e2 on c.spouseconstituentdimid = e2.constituentdimid
left join #sitegiving_pivot top3 on #ids.constituentdimid = top3.constituentdimid
left join #sitegiving_pivot top3spouse on c.spouseconstituentdimid = top3spouse.constituentdimid
left join #camp_sitegiving_pivot top3camp on #ids.constituentdimid = top3camp.constituentdimid
left join #camp_sitegiving_pivot top3campspouse on c.spouseconstituentdimid = top3campspouse.constituentdimid
left join #last_interactions lc on #ids.constituentdimid = lc.constituentdimid and lc.seq = 1
left join #last_interactions lv on #ids.constituentdimid = lv.constituentdimid and lv.visitseq = 1
left join #board b1 on b1.constituentdimid = #ids.constituentdimid and b1.keyname = 'unc board of trustees'
left join #board b2 on b2.constituentdimid = #ids.constituentdimid and b2.keyname = 'chancellor’s philanthropic council'
left join #board b3 on b3.constituentdimid = #ids.constituentdimid and b3.keyname = 'chancellor’s philanthropic council emeriti'
left join #board b4 on b4.constituentdimid = #ids.constituentdimid and b4.keyname = 'campaign for carolina steering committee'
left join #board b5 on b5.constituentdimid = #ids.constituentdimid and b5.keyname = 'carolina 1st campaign cabinet members 98-99'

left join 
(
select 
constituentdimid, sum(gifts + pledges + irrevocable + revocable + grants) as amount
from bbinfinity_rpt_bbdw.bbdw.fact_usr_unc_aggregatecampaigngiving_constituent a
group by constituentdimid
) agglegal on agglegal.constituentdimid = d.constituentdimid

left join 
(
select constituentdimid, sum(gifts + pledges + irrevocable + revocable + grants) as amount
from bbinfinity_rpt_bbdw.bbdw.fact_usr_unc_aggregatecampaigngiving_constituent 
group by constituentdimid
) agglegalspouse on agglegalspouse.constituentdimid = c.spouseconstituentdimid

outer apply
(
select top 1 
fsc.constituentsystemid, dsc.specialcode 

from 
bbinfinity_rpt_bbdw.bbdw.fact_usr_unc_constituentspecialcode fsc 
join bbinfinity_rpt_bbdw.bbdw.dim_usr_unc_specialcode dsc on fsc.specialcodedimid = dsc.specialcodedimid
join #ids on #ids.constituentsystemid = fsc.constituentsystemid

where
dsc.specialcode in ('PGP') and d.constituentsystemid = fsc.constituentsystemid
) pgp

outer apply
(
select top 1 
fsc.constituentsystemid, dsc.specialcode 

from 
bbinfinity_rpt_bbdw.bbdw.fact_usr_unc_constituentspecialcode fsc 
join bbinfinity_rpt_bbdw.bbdw.dim_usr_unc_specialcode dsc on fsc.specialcodedimid = dsc.specialcodedimid
join #ids on #ids.constituentsystemid = fsc.constituentsystemid

where
dsc.specialcode  in ('1PG') and d.constituentsystemid = fsc.constituentsystemid
) ogp

outer apply
(
select top 1 
fsc.constituentsystemid, dsc.specialcode 

from 
bbinfinity_rpt_bbdw.bbdw.fact_usr_unc_constituentspecialcode fsc 
join bbinfinity_rpt_bbdw.bbdw.dim_usr_unc_specialcode dsc on fsc.specialcodedimid = dsc.specialcodedimid
join #ids on #ids.constituentsystemid = fsc.constituentsystemid

where
dsc.specialcode  in ('FPG') and d.constituentsystemid = fsc.constituentsystemid
) fpg

outer apply
(
select top 1 
fsc.constituentsystemid, dsc.specialcode 

from 
bbinfinity_rpt_bbdw.bbdw.fact_usr_unc_constituentspecialcode fsc 
join bbinfinity_rpt_bbdw.bbdw.dim_usr_unc_specialcode dsc on fsc.specialcodedimid = dsc.specialcodedimid
join #ids on #ids.constituentsystemid = fsc.constituentsystemid

where
dsc.specialcode  in ('EWP') and d.constituentsystemid = fsc.constituentsystemid
) ewp

outer apply
(
select top 1 
fsc.constituentsystemid, dsc.specialcode 

from 
bbinfinity_rpt_bbdw.bbdw.fact_usr_unc_constituentspecialcode fsc 
join bbinfinity_rpt_bbdw.bbdw.dim_usr_unc_specialcode dsc on fsc.specialcodedimid = dsc.specialcodedimid
join #ids on #ids.constituentsystemid = fsc.constituentsystemid

where
dsc.specialcode  in ('1WP') and d.constituentsystemid = fsc.constituentsystemid
) owp


select 
*, 
bonus_ethnicity_score +
bonus_specialcode_score +
bonus_mngr_score +
bonus_ss_score +
bonus_giving_score as bonus_score,

bonus_ethnicity_score +
bonus_specialcode_score +
bonus_mngr_score +
bonus_ss_score +
bonus_giving_score +
affinity_score +
capacity_score as total_score

from 
#results

order by
bonus_ethnicity_score +
bonus_specialcode_score +
bonus_mngr_score +
bonus_ss_score +
bonus_giving_score +
affinity_score +
capacity_score desc

end