

-- current season for season ticket holder section
declare @fbseason as nvarchar(4) = (select top 1 season from usr_unc_ssb_order where item = 'fs' order by season desc);
declare @bbseason as nvarchar(4) = (select top 1 season from usr_unc_ssb_order where item = 'bs' order by season desc);
declare @wbbseason as nvarchar(4) = (select top 1  season from usr_unc_ssb_order where item = 'ws' and season like 'wb%' order by season desc);
declare @bseason as nvarchar(4) = (select top 1 season from usr_unc_ssb_order where item = 'bbs' order by season desc);

-- board types for board section
declare @boardtype table
(
boardname varchar(200),
boardtype varchar(200)
);

insert into @boardtype
values
('Carolina Native Alumni Circle', 'Alumni Assoc Boards'),
('KFBS Alumni Council', 'Alumni Assoc Boards'),
('Law Alumni Association Board of Directors', 'Alumni Assoc Boards'),
('Medical Alumni Council', 'Alumni Assoc Boards'),
('MPA Alumni Association Board', 'Alumni Assoc Boards'),
('Pharmacy Alumni Association Board of Directors', 'Alumni Assoc Boards'),
('School of Education Alumni Council', 'Alumni Assoc Boards'),
('School of Social Work Alumni Council', 'Alumni Assoc Boards'),
('SON Alumni Association Board of Directors', 'Alumni Assoc Boards'),
('Board of Governors', 'Board of Governors'),
('UNC Board of Trustees', 'Board of Trustees'),
('Children''s Hospital Board of Visitors', 'Board of Visitors'),
('Eating Disorders Board of Visitors', 'Board of Visitors'),
('Health Science Library Board of Visitors', 'Board of Visitors'),
('Honorary Board of Visitors', 'Board of Visitors'),
('Institute for the Environment Board of Visitors', 'Board of Visitors'),
('Lineberger Cancer Center Board of Visitors', 'Board of Visitors'),
('Pharmacy Board of Visitors', 'Board of Visitors'),
('Psychiatry Board of Visitors', 'Board of Visitors'),
('SILS Board of Visitors', 'Board of Visitors'),
('UNC Board of Visitors', 'Board of Visitors');

-- greek types for involvement section
declare @greektype table
(
greekname varchar(200),
greektype varchar(200)
);

insert into @greektype
values
('Alpha Delta Pi','three point'),
('Alpha Tau Omega Fraternity-ATO','three point'),
('Chi Omega Sorority-CO','three point'),
('Delta Kappa Epsilon Fraternity-DKE','three point'),
('Phi Delta Theta Fraternity-PDT','three point'),
('Phi Gamma Delta Fraternity-PGD','three point'),
('Pi Beta Phi Sorority-PBP','three point'),
('Zeta Psi Fraternity-ZP','three point'),
('Alpha Chi Omega Sorority-ACO','other'),
('Alpha Epsilon Pi Fraternity-AEP','other'),
('Alpha Gamma Delta Sorority-AGD','other'),
('Alpha Kappa Alpha Sorority-AKA','other'),
('Alpha Kappa Delta Phi Sorority=AKDP','other'),
('Alpha Kappa Psi Fraternity-AKP','other'),
('Alpha Phi Alpha Fraternity-APA','other'),
('Alpha Phi','other'),
('Alpha Pi Omega-APOS','other'),
('Beta Theta Pi Fraternity-BTF','other'),
('Beta Upsilon Chi','other'),
('Chi Phi Fraternity-CP','other'),
('Chi Psi Fraternity-CPF','other'),
('Delta Chi','other'),
('Delta Delta Delta Sorority-DDD','other'),
('Delta Phi Omega Sorority, Incorporated','other'),
('Delta Phi Epsilon Sorority-DPE','other'),
('Delta Sigma Iota Fraternity, Incorporated','other'),
('Delta Sigma Iota-DSI','other'),
('Delta Sigma Phi','other'),
('Delta Sigma Phi Fraternity-DSF','other'),
('Delta Sigma Theta Sorority-DST','other'),
('Delta Tau Delta Fraternity-DTD','other'),
('Delta Upsilon Fraternity-DUF','other'),
('Delta Zeta Sorority-DZ','other'),
('Kappa Alpha Order Fraternity-KAO','other'),
('Kappa Alpha Psi Fraternity-KAP','other'),
('Kappa Alpha Theta Sorority-KAT','other'),
('Kappa Delta Sorority','other'),
('Kappa Kappa Gamma Sorority-KKG','other'),
('Kappa Phi Lambda Sorority-KPL','other'),
('Kappa Sigma Fraternity-KS','other'),
('La Unidad Latina, Lambda Upsilon Lambda Fraternity, Incorporated','other'),
('Lambda Chi Alpha Fraternity-LCA','other'),
('Lambda Phi Epsilon Fraternity-LPE','other'),
('Lambda Pi Chi Sorority-LPC','other'),
('Omega Phi Beta-OPB','other'),
('Omega Psi Phi Fraternity-OPP','other'),
('Phi Beta Chi-PBC','other'),
('Phi Beta Sigma Fraternity','other'),
('Phi Beta Sigma Fraternity-PBS','other'),
('Phi Kappa Sigma Fraternity-PKS','other'),
('Phi Kappa Tau Fraternity-PKT','other'),
('Phi Mu Sorority-PM','other'),
('Phi Sigma Nu Fraternity-NSP','other'),
('Pi Alph Phi Fraternity-FPA','other'),
('Pi Kappa Alpha Fraternity-PKA','other'),
('Pi Kappa Phi Fraternity-PKP','other'),
('Pi Lambda Phi Fraternity-PLP','other'),
('Psi Sigma Phi Fraternity-FPS','other'),
('Sigma Alpha Epsilon Fraternity-SAE','other'),
('Sigma Chi Fraternity-SCF','other'),
('Sigma Gamma Rho Sorority-SGR','other'),
('Sigma Nu Fraternity-SN','other'),
('Sigma Phi','other'),
('Sigma Phi Epsilon Fraternity-SPE','other'),
('Sigma Phi Society Fraternity-SPS','other'),
('Sigma Rho Lambda Sorority-SRL','other'),
('Sigma Sigma Sigma Sorority-SSS','other'),
('Tau Epsilon Phi Fraternity-TEP','other'),
('St. Anthony Hall','other'),
('Tau Kappa Epsilon Fraternity-TKE','other'),
('Theta Nu Xi Multicultural Sorority, Incorporated','other'),
('Zeta Beta Tau Fraternity-ZBT','other'),
('Zeta Phi Beta Sorority-ZPB','other'),
('Zeta Tau Alpha Sorority-ZTA','other'),
('Alpha Sigma Phi','other');


-- 1) unc degree - 2.5 point maximum
with relation as
(
select
r.relationshipconstituentid,
c1.lookupid,
c1.firstname,
c1.keyname,
rtc.description relationship_type,
r.reciprocalconstituentid,
c2.lookupid recip_lookupid,
c2.firstname recip_firstname,
c2.keyname recip_keyname

from 
relationship r,
relationshiptypecode rtc,
constituent c1,
constituent c2

where
r.relationshiptypecodeid = rtc.id
and r.relationshipconstituentid = c1.id
and r.reciprocalconstituentid = c2.id
  and rtc.description in 
  ('Aunt', 'Brother', 'Child', 'Cousin', 'Daughter', 'Family Member', 'Foster Child', 'Foster Parent', 'Godchild', 'Goddaughter', 'Godparent', 'Godson,', 
  'Grandchild', 'Granddaughter', 'Grandparent', 'Grandson', 'Great-aunt', 'Great-grandchild', 'Great-granddaughter', 'Great-grandparent', 'Great-grandson', 'Great-nephew', 
  'Great-niece', 'Great-uncle', 'Half-brother', 'Guardian', 'Half-sibling', 'Half-sister', 'Nephew', 'Niece', 'Parent', 'Sibling', 'Sister', 'Uncle', 'Step-sibling', 'Step-parent', 
  'Step-son', 'Step-sister', 'Step-daughter', 'Step-child', 'Step-brother', 'Son')
),

degree as
(
select
x.constituentid,
sum(x.degree_undergrad_ind) as degree_undergrad_ind,
sum(x.degree_grad_ind) as degree_grad_ind

from
(
  select
  eh.constituentid,
  case 
  when edc.description like 'bachelor%' then 1
  when edc.description like 'bs %' then 1
  when edc.description like 'ba %' then 1
  else 0 end as degree_undergrad_ind,
  case 
  when edc.description like 'master%' then 1
  when edc.description like 'ms %' then 1
  when edc.description like 'ma %' then 1
  when edc.description like 'doctor %' then 1
  when edc.description like 'doctorate %' then 1
  when edc.description like 'juris doctor%' then 1
  else 0 end as degree_grad_ind

  from
  dbo.educationalhistory eh, dbo.educationaldegreecode edc

  where
  eh.educationaldegreecodeid = edc.id and eh.educationalinstitutionid = '67673e25-caf3-4d45-a4d2-ef633736b7d0'
) x

group by x.constituentid
),

degree_relation as
(
select
d1.constituentid,
d1.degree_undergrad_ind,
d1.degree_grad_ind,
r.lookupid,
r.firstname,
r.keyname,
r.relationship_type,
r.reciprocalconstituentid,
r.recip_lookupid,
r.recip_firstname,
r.recip_keyname,
d2.degree_undergrad_ind as recip_degree_undergrad_ind,
d2.degree_grad_ind as recip_degree_grad_ind,
case 
when d1.degree_undergrad_ind > 0 and d1.degree_grad_ind > 0 then 2.5
when (d1.degree_undergrad_ind + d1.degree_grad_ind) > 0 and (d2.degree_undergrad_ind + d2.degree_grad_ind) > 0 and r.relationship_type in ('Spouse', 'Spouse/Partner') then 2.5
when d1.degree_undergrad_ind > 0 and d1.degree_grad_ind = 0 then 1.5
when (d1.degree_undergrad_ind + d1.degree_grad_ind) > 0 and (d2.degree_undergrad_ind + d2.degree_grad_ind) > 0 and r.relationship_type not in ('Spouse', 'Spouse/Partner') then 1.5
when d1.degree_grad_ind > 0 and d1.degree_undergrad_ind = 0 then 1 
when d1.degree_undergrad_ind = 0 and d1.degree_grad_ind = 0 and (d2.degree_undergrad_ind + d2.degree_grad_ind) > 0 then .5 
else 0 end as score_degree

from
degree d1 
left join relation r on d1.constituentid = r.relationshipconstituentid
left join degree d2 on r.reciprocalconstituentid = d2.constituentid
),

degree_score as (select constituentid, max(score_degree) as degree_score from degree_relation group by constituentid),


-- 2) parent of unc student - 2.5 point maximum
parent_score as
(
select
x.relationshipconstituentid constituentid,
case when x.parent_current_ind = 1 then 2.5 when x.parent_past_ind = 1 then 1 else 0 end as parent_score

from
(
select
r.relationshipconstituentid,
max(case when eh.constituencystatus in ('currently attending') then 1 else 0 end) as parent_current_ind,
max(case when eh.constituencystatus in ('graduated') then 1 else 0 end) as parent_past_ind

from
dbo.relationship r, dbo.educationalhistory eh, dbo.relationshiptypecode rtc

where
r.reciprocalconstituentid = eh.constituentid 
and r.relationshiptypecodeid = rtc.id
  and eh.educationalinstitutionid = '67673e25-caf3-4d45-a4d2-ef633736b7d0'
  and rtc.description in ('parent', 'guardian')

group by
r.relationshipconstituentid
) x
),


-- 3) unc boards - 2.5 point maximum
board_detail as
(
select distinct
c1.id constituentid, c1.lookupid, c1.firstname + ' ' + c1.keyname constituent_name, c2.lookupid groupid, c2.keyname group_name, gt.name grouptype,
case when c2.keyname = 'Chancellor’s Philanthropic Council' then 1 else 0 end as board_philanthropic_ind,
case when bd.boardtype = 'Alumni Assoc Boards' then 1 else 0 end as board_alumni_ind,
case when bd.boardtype = 'Board of Governors' then 1 else 0 end as board_governors_ind,
case when bd.boardtype = 'Board of Trustees' then 1 else 0 end as board_trustees_ind,
case when bd.boardtype = 'Board of Visitors' then 1 else 0 end as board_visitors_ind,
1 as board_other_ind

from
dbo.constituent c1
join dbo.groupmember gm on gm.memberid = c1.id
join dbo.constituent c2 on gm.groupid = c2.id
join dbo.groupdata gd on gm.groupid = gd.id
join dbo.grouptype gt on gt.id = gd.grouptypeid
join (select distinct groupmemberid from dbo.groupmemberdaterange where dateto is null or dateto >= '01-01-2015') dr on gm.id = dr.groupmemberid
left join @boardtype bd on c2.keyname = bd.boardname

where
gt.name in ('board', 'fundraising committee', 'committee', 'reunion committee')
),

board_summary as
(
select 
constituentid,
sum(board_philanthropic_ind) as board_philanthropic_ind,
sum(board_alumni_ind) as board_alumni_ind,
sum(board_governors_ind) as board_governors_ind,
sum(board_trustees_ind) as board_trustees_ind,
sum(board_visitors_ind) as board_visitors_ind,
sum(board_other_ind) as board_other_ind

from 
board_detail

group by
constituentid
),

board_score as
(
select
constituentid,
case 
when board_philanthropic_ind >= 1 then 2.5
when board_alumni_ind > 1 then 2.5
when board_visitors_ind >= 1 then 1.75
when board_trustees_ind >= 1 then 1.5
when board_alumni_ind = 1 then 1.5
when board_governors_ind >= 0 then .5
when board_other_ind >= 0 then .25
else 0 end as board_score 

from
board_summary
),

-- 4) lifetime giving by capacity rating - 5 point maximum
giving_unc_rating as
(
select 
p.id constituentid, 
max(case
when psc.description like 'A__-%' then 5000000
when psc.description like 'AA__-%' then 10000000
when psc.description like 'AAA__-%' then 25000000
when psc.description like 'AAAA__-%' then 50000000
when psc.description like 'AAAAA__-%' then 100000000
when psc.description like 'B__-%' then 1000000
when psc.description like 'C__-%' then 500000
when psc.description like 'D__-%' then 100000
when psc.description like 'E__-%' then 25000
when psc.description like 'XAAAAA%' then 100000000
when psc.description like 'XAAAA%' then 50000000
when psc.description like 'XAAA%' then 25000000
when psc.description like 'XAA%' then 10000000
when psc.description like 'XA%' then 5000000
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

giving_gga_rating as
(
select
r.modelingandpropensityid constituentid, 
max(case
when c.description like '1%' then 10000000
when c.description like '2%' then 1000000
when c.description like '3%' then 250000
when c.description like '4%' then 100000
when c.description like '5%' then 25000
when c.description like '6%' then 10000
else 0 end) as gga_capacity_amt

from 
dbo.attribute211d1d9e300f4e438330b676977ec212 r, dbo.ggagiftcapacityratingcode c

where 
c.id = r.ggagiftcapacityratingcodeid

group by
modelingandpropensityid
),

giving_donorscape_rating as
(
select
constituentid, 
max(case
when ExactNearGiftCapacityRating like '1%' then 10000000
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

giving_alltime as
(
select 
a.constituentid, sum(a.amount) giving_amt

from 
usr_unc_constituentgivingaggregatebase a

where 
a.amount > 0 and a.isedfoundation = 0 and a.ispledgepayment = 0 and a.isgrant = 0 and a.isrecognized = 1

group by 
a.constituentid
),

giving_capacity as
(
select
a.constituentid, 
a.giving_amt, 
case when u.unc_rating_amt is null then d.donorscape_capacity_amt else u.unc_rating_amt end as denominator

from 
giving_alltime a
left join giving_unc_rating u on a.constituentid = u.constituentid
left join giving_gga_rating g on a.constituentid = g.constituentid
left join giving_donorscape_rating d on a.constituentid = d.constituentid
),

giving_score as
(
select
constituentid,
case 
when denominator = 0 then .5
when giving_amt / denominator >= .8 then 4.5
when giving_amt / denominator >= .6 then 3.5
when giving_amt / denominator >= .4 then 2.5
when giving_amt / denominator >= .2 then 1.5
else .5 end as giving_score

from 
giving_capacity
),

-- 5) recency + frequency combined - 2.5 point maximum
rfm_score as
(
select
constituentid, rec, freq, rfm,
case
when isnull(rec, 0) + isnull(freq, 0)  >= 40 then 2.5
when isnull(rec, 0) + isnull(freq, 0)  >= 30 then 2
when isnull(rec, 0) + isnull(freq, 0)  >= 20 then 1.5
when isnull(rec, 0) + isnull(freq, 0)  >= 10 then 1
else 0 end as rfm_score

from
bbinfinity_rpt_bbdw.dbo.usr_unc_pmra_rfmshortfile
),

-- 6) educational involvement - 3.5 point maximum
involvement_detail as
(
select
ei.constituentid,
case when eitc.description in ('athletics') then 1 else 0 end as involve_athletics_ind, 
case when eitc.description in ('educational award') and ein.name not in ('Study Abroad-ABR') then 1 else 0 end as involve_ed_award_ind, 
case when eitc.description in ('educational award') and ein.name in ('Study Abroad-ABR') then 1 else 0 end as involve_study_abroad_ind, 
case when eitc.description in ('greek', 'greek organization') and g.greektype = 'three point' then 1 else 0 end as involve_greek_three_ind, 
case when eitc.description in ('greek', 'greek organization') and g.greektype = 'other' then 1 else 0 end as involve_greek_other_ind, 
case when eitc.description in ('intramurals') then 1 else 0 end as involve_intramural_ind, 
case when eitc.description in ('student organization', 'student government') then 1 else 0 end as involve_student_orgn_ind, 
case when eitc.description in ('study area') then 1 else 0 end as involve_study_area_ind,
case when eitc.description in ('volunteer') then 1 else 0 end as involve_volunteer_ind

from 
dbo.educationalinvolvement ei
join educationalinvolvementtypecode eitc on ei.educationalinvolvementtypecodeid = eitc.id
join educationalinvolvementname ein on ei.educationalinvolvementnameid = ein.id
left join @greektype g on ein.name = g.greekname
), 

involvement_summary as
(
select
constituentid,
max(involve_athletics_ind) as involve_athletics_ind,
max(involve_ed_award_ind) as involve_ed_award_ind,
max(involve_study_abroad_ind) as involve_study_abroad_ind,
max(involve_greek_three_ind) as involve_greek_three_ind,
max(involve_greek_other_ind) as involve_greek_other_ind,
max(involve_intramural_ind) as involve_intramural_ind,
max(involve_student_orgn_ind) as involve_student_orgn_ind,
max(involve_study_area_ind) as involve_study_area_ind,
max(involve_volunteer_ind) as involve_volunteer_ind,
max(involve_athletics_ind) + max(involve_ed_award_ind) + max(involve_study_abroad_ind) +
max(involve_greek_three_ind) + max(involve_greek_other_ind) + max(involve_intramural_ind) + max(involve_student_orgn_ind) +
max(involve_study_area_ind) + max(involve_volunteer_ind) as involve_multiple_ind 

from
involvement_detail

group by
constituentid
),

involvement_score as
(
select
constituentid,
case 
when involve_multiple_ind > 1 then 3.5
when involve_greek_three_ind = 1 then 3
when involve_greek_other_ind = 1 then 2.5
when involve_ed_award_ind = 1 then 2.5
when involve_athletics_ind = 1 then 1.5
when involve_intramural_ind = 1 then 1
when involve_study_area_ind = 1 then 1
when involve_student_orgn_ind = 1 then 1
when involve_volunteer_ind = 1 then 1
when involve_study_abroad_ind = 1 then 1
else 0 end as involvement_score

from 
involvement_summary
),

-- 7) event attendance - 4 point maximum
event_score as
(
select 
constituentid,
case 
when event_cnt >= 4 then 4
when event_cnt >= 3 then 3
when event_cnt >= 2 then 2
when event_cnt = 1 then .5
else 0 end as event_score

from
(
select r.constituentid, count(*) event_cnt 
from registrant r join event e on r.eventid = e.id
where r.attended = 1 and e.startdate >= '01-01-2015'
group by r.constituentid) x
),

-- 8) alumni club membership - 1.5 point maximum
gaa_detail as
(
select
m.constituentid, mp.name, ms.expirationdate,
case when mp.name like 'GAA: Lifetime%' and isnull(ms.expirationdate, '9999-12-31') >= getdate() then 1 else 0 end as gaa_lifetime_ind,
case when mp.name like 'GAA: Annual%' and isnull(ms.expirationdate, '9999-12-31')  >= getdate() then 1 else 0 end as gaa_annual_ind,
case when ms.expirationdate < getdate() then 1 else 0 end as gaa_lapsed_ind

from
member m
join membership ms on m.membershipid = ms.id
join membershipprogram mp on ms.membershipprogramid = mp.id

where
mp.name like 'GAA: Annual%' or mp.name like 'GAA: Lifetime%'
),

gaa_summary as
(
select constituentid, max(gaa_lifetime_ind) as gaa_lifetime_ind, max(gaa_annual_ind) as gaa_annual_ind, max(gaa_lapsed_ind) as gaa_lapsed_ind
from gaa_detail
group by constituentid
),

gaa_score as
(
select
constituentid,
case 
when gaa_lifetime_ind = 1 then 1.5
when gaa_annual_ind = 1 then 1
when gaa_lapsed_ind = 1 then .5
else 0 end as gaa_score

from
gaa_summary
),

-- 9) season ticket holder - 1.5 point maximum
ticket_detail as
(
select distinct
i.constituentid,
o.season,
o.item,
case 
when o.season = @fbseason then 1.5
when o.season = @bbseason then 1.5
when o.season = @wbbseason then 1.5
when o.season = @bseason then 1
else 0 end as ticket_score

from
usr_unc_constituentpaciolanid i join usr_unc_ssb_order o on i.paciolanid = o.customer

where
o.item in ('FS', 'BS', 'WS', 'BBS')
and o.season in (@fbseason, @bbseason, @wbbseason, @bseason)
),

ticket_score as
(select constituentid, max(ticket_score) as ticket_score from ticket_detail group by constituentid),

-- merge scores
scores as
(
select distinct 
c.id CONSTITUENTSYSTEMID,
c.lookupid CONSTITUENTLOOKUPID,
d.constituentdimid CONSTITUENTDIMID,
c.keyname LASTNAME,
c.firstname + ' ' + c.keyname FULLNAME,
d.primaryconstituency PRIMARYCONSTITUENCY,
c.gender GENDER,
d.isalumnus ISALUMNUS,
c.isinactive ISINACTIVE,
c.isorganization ISORGANIZATION,
ch.householdid HOUSEHOLDID,
d.spousefullname SPOUSEFULLNAME,
isnull(ds.degree_score, 0)*4 as DEGREE_SCORE,
isnull(ps.parent_score, 0)*4 as PARENT_SCORE,
isnull(bs.board_score, 0)*4 as BOARD_SCORE,
isnull(gs.giving_score, 0)*4 as GIVING_SCORE,
isnull(rs.rfm_score, 0)*4 as RFM_SCORE,
isnull(i.involvement_score, 0)*4 as INVOLVEMENT_SCORE,
isnull(es.event_score, 0)*4 as EVENT_SCORE,
isnull(gaa.gaa_score, 0)*4 as GAA_SCORE,
isnull(ts.ticket_score, 0)*4 as TICKET_SCORE,
(isnull(ds.degree_score, 0) + isnull(ps.parent_score, 0) + isnull(bs.board_score, 0) + isnull(gs.giving_score, 0) + isnull(rs.rfm_score, 0) + 
isnull(i.involvement_score, 0) + isnull(es.event_score, 0) + isnull(gaa.gaa_score, 0) + isnull(ts.ticket_score, 0)) * 4 as TOTAL_SCORE

from
dbo.constituent c
join bbinfinity_rpt_bbdw.bbdw.fact_usr_unc_demographic d on d.constituentsystemid = c.id
left join deceasedconstituent dc on c.id = dc.id
left join constituenthousehold ch on c.id = ch.id
left join degree_score ds on c.id = ds.constituentid
left join parent_score ps on c.id = ps.constituentid
left join board_score bs on c.id = bs.constituentid
left join giving_score gs on c.id = gs.constituentid
left join rfm_score rs on c.id = rs.constituentid
left join involvement_score i on c.id = i.constituentid
left join event_score es on c.id = es.constituentid
left join gaa_score gaa on c.id = gaa.constituentid
left join ticket_score ts on c.id = ts.constituentid

where
d.isconstituent = 1
and d.isorganization = 0 
and d.isgroup = 0
and d.isinactive = 0
and dc.id is null
)

select
s1.*, isnull(sp.total_score, 0) as SPOUSE_TOTAL_SCORE, 
case when s1.total_score > isnull(sp.total_score, 0) then s1.total_score else isnull(sp.total_score, 0)  end as FINAL_TOTAL_SCORE

from 
scores s1 
--left join scores s2 on s1.householdid = s2.householdid and s1.constituentsystemid <> s2.constituentsystemid

outer apply 
(
select top 1 constituentlookupid, fullname, total_score
from scores s2
where s1.householdid = s2.householdid and s1.constituentsystemid <> s2.constituentsystemid
) sp