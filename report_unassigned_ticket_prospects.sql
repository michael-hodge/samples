
create procedure [dbo].[USR_UNC_USP_UNASSIGNED_TICKET_PROSPECT_REPORT] 

as begin

select 
r.relationshipconstituentid as constituentid, r.id, history.classof

into #children

from 
relationship r
join relationshiptypecode on r.relationshiptypecodeid = relationshiptypecode.id
join dbo.educationalhistory as history on r.reciprocalconstituentid = history.constituentid				
left join dbo.educationalhistorystatus educationstatus on history.educationalhistorystatusid = educationstatus.id

where 
relationshiptypecode.description in ('foster parent', 'god-parent', 'parent', 'step parent')
and history.educationalinstitutionid = '67673e25-caf3-4d45-a4d2-ef633736b7d0'
and history.classof > 0
and educationstatus.abbreviation in ('a','g','i');

create nonclustered index ix_children on #children (constituentid);


select distinct
c1.constituentid,
stuff
((
select distinct cast(',' as varchar(60)) + rtrim(ltrim(c2.classof))
from #children c2
where c1.constituentid = c2.constituentid
for xml path('')
), 1, 1, '') as childclass

into #childclass

from 
#children c1;

create nonclustered index ix_childclass on #childclass (constituentid);


select 
cdr.constituentid

into #staff

from
constituencydaterange cdr
join constituencydefinition cd on cdr.constituencydefinitionid = cd.id and cd.description = 'faculty/staff'
join usr_unc_constituencyorder co on cdr.constituencydefinitionid = co.constituencycodeid

where 
dateto is null;

create nonclustered index ix_staff on #staff (constituentid);


select constituentid, leadgroup, row_number() over(partition by constituentid order by dateadded desc) as seq
into #leadgroup
from usr_unc_constituentleadgroup lg
where lg.enddate is null
group by constituentid, leadgroup, dateadded;

create nonclustered index ix_leadgroup on #leadgroup (constituentid, seq);


select i.constituentid, i.comment, row_number() over(partition by i.constituentid order by i.dateadded desc) as seq
into #lastinteraction
from interaction i
join constituency cc on i.fundraiserid = cc.constituentid
join constituencycode ccc on cc.constituencycodeid = ccc.id and ccc.description = 'ticket rep';

create nonclustered index ix_lastinteraction on #lastinteraction (constituentid, seq);


select distinct 
cc.id as constituentid, t.season, t.i_pt, t.item

into #tickets

from 
usr_unc_ssb_customer a
join constituent cc on a.ssid = cc.id 
join usr_unc_ssb_order t on a.ssb_crmsystem_contact_id = t.ssb_crmsystem_contact_id

where 
a.sourcesystem = 'blackbaud'
and t.i_pt is not null
and t.i_pt <> ''
and (t.season like 'fb%' or t.season like 'mb%');

create nonclustered index ix_tickets on #tickets (constituentid, season);


select distinct
t1.constituentid,
stuff
((
select distinct cast(',' as varchar(150)) + rtrim(ltrim(t2.i_pt))
from #tickets t2
where t1.constituentid = t2.constituentid and t2.season like 'fb%' 
for xml path('')
), 1, 1, '') as fbticket

into #fbticket

from 
#tickets t1

where 
t1.season like 'fb%';;

create nonclustered index ix_fbticket on #fbticket (constituentid);


select distinct
t1.constituentid,
stuff
((
select distinct cast(',' as varchar(150)) + rtrim(ltrim(t2.i_pt))
from #tickets t2
where t1.constituentid = t2.constituentid and t2.season like 'mb%' 
for xml path('')
), 1, 1, '') as bbticket

into #bbticket

from 
#tickets t1

where 
t1.season like 'mb%';

create nonclustered index ix_bbticket on #bbticket (constituentid);


select distinct
t1.constituentid,
stuff
((
select distinct cast(',' as varchar(150)) + rtrim(ltrim(t2.item))
from #tickets t2
where t1.constituentid = t2.constituentid and t2.season like 'fb%' 
for xml path('')
), 1, 1, '') as fbitem

into #fbitem

from 
#tickets t1

where 
t1.season like 'fb%';

create nonclustered index ix_fbitem on #fbitem (constituentid);


select distinct
t1.constituentid,
stuff
((
select distinct cast(',' as varchar(150)) + rtrim(ltrim(t2.item))
from #tickets t2
where t1.constituentid = t2.constituentid and t2.season like 'mb%' 
for xml path('')
), 1, 1, '') as bbitem

into #bbitem

from 
#tickets t1

where 
t1.season like 'mb%';

create nonclustered index ix_bbitem on #bbitem (constituentid);


select distinct
p.constituentid, c.emailprimary, row_number() over(partition by p.constituentid order by c.emailprimary) as seq

into #email

from
usr_unc_ssb_customer c 
join usr_unc_constituentpaciolanid p on c.ssid = p.paciolanid and c.sourcesystem = 'pac'

where
c.emailprimary <> '' and c.donotemail = 0;

create nonclustered index ix_email on #email (constituentid, seq);


select distinct top 100000
c.lookupid,
c.id, 
c.name,
--cp.paciolanid,
p.number phone,
ptc.description phonetype,
e.emailaddress,
email.emailprimary as pacemail,
a.addressblock,
a.city,
s.description state, 
a.postcode, 
r.employer, 
r.jobtitle, 
r.spousefullname,
case when r.primaryclass = 0 then null
else r.primaryclass end as primaryclass,
cc.fbcurrentyr,
cc.fblybntyr,
cc.fbsybntyr,
cc.bbcurrentyr,
cc.bblybntyr,
cc.bbsybntyr,
rg.lygivinglevel,
rg.cygivinglevel,
case when st.constituentid is not null then 'yes'
else 'no' end as staff,
fb.fbticket,
bb.bbticket,
fbi.fbitem,
bbi.bbitem,
i.comment,
lg.leadgroup,

(
select distinct cast(stuff((select ',' + paciolanid 
from dbo.usr_unc_constituentpaciolanid  
where constituentid = c.id
for xml path('')), 1, 1, '')  as varchar(150))
) pacids,
case when ch.constituentid is not null then 'yes'
else 'no' end as parent,
ch.childclass as studentclassyear,
'https://davie-crm.dev.unc.edu/bbappfx/webui/webshellpage.aspx?databasename=BBInfinity#pageType=p&pageId=88159265-2b7e-4c7b-82a2-119d01ecd40f&tabId=2e862877-34e8-47a1-8545-2a5ccbee876d&recordId=' + cast(c.id as varchar(36)) as url

from 
usr_unc_ssb_customer aa
join constituent c on aa.ssid=c.id
left join phone p on p.constituentid = c.id and p.isprimary = 1
left join phonetypecode ptc on ptc.id = p.phonetypecodeid
left join address a on a.constituentid = c.id and a.addresstypecodeid = '7cd72b86-452b-4360-a77c-d24fe3d5081d' and a.historicalenddate is null --home mailing
left join state s on a.stateid = s.id
left join emailaddress e on e.constituentid = c.id and e.isprimary = 1
left join bbinfinity_rpt_bbdw.bbdw.fact_usr_unc_demographic r on r.constituentsystemid = c.id
left join prospectteam pt on pt.prospectid = c.id and prospectteamrolecodeid = '9c759793-62f5-47c5-af34-abf327d4a6b3'
join usr_unc_constituentpaciolanid cp on cp.constituentid = c.id
left join usr_unc_uncticketsummary cc on cc.constituentid = c.id
left join usr_unc_constituent_ext_gp gp on gp.id = c.id
left join usr_unc_constituentrcgivinglevel rg on rg.constituentid = c.id

left join #childclass ch on c.id = ch.constituentid
left join #staff st on c.id = st.constituentid
left join #leadgroup lg on c.id = lg.constituentid and lg.seq = 1
left join #lastinteraction i on c.id = i.constituentid and i.seq = 1
left join #email email on c.id = email.constituentid and email.seq = 1
left join #fbticket fb on fb.constituentid = c.id
left join #bbticket bb on bb.constituentid = c.id
left join #fbitem fbi on fbi.constituentid = c.id
left join #bbitem bbi on bbi.constituentid = c.id

where 
aa.sourcesystem='blackbaud'
and pt.id is null
and len(aa.ssid) = 36

and cp.paciolanid in
(
select distinct a2.ssid 
from dbo.usr_unc_ssb_order t join usr_unc_ssb_customer a2 on t.customer = a2.ssid 
)

and c.id not in
(
select id from deceasedconstituent
where deceasedconfirmationcode > 0
)

and (gp.id is null or gp.patientonly = 0)

end 