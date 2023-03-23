


-- assign fiscal year variables
declare @CurrentFYID smallint;
declare @CurrentFY smallint;

set @CurrentFYID = (select y.yearid fiscalyear from dbo.glfiscalperiod p join glfiscalyear y on p.glfiscalyearid = y.id where getdate() between p.startdate and p.enddate);
set @CurrentFY = 2000 + @CurrentFYID;


--------------------------------------------------------------------------------
-- get years
--------------------------------------------------------------------------------
select top 2023 row_number() over(order by id) yr
into #years
from constituent

delete from #years where yr < 1980

--------------------------------------------------------------------------------
-- fiscal year dates
--------------------------------------------------------------------------------
select
convert(integer, substring(y.description, 6, 4)) as fiscalyear,
min(p.startdate) startdate,
max(p.enddate) enddate

into
#fiscalyear_dates

from
glfiscalyear y join glfiscalperiod p on y.id = p.glfiscalyearid

group by 
convert(integer, substring(y.description, 6, 4))

order by 
convert(integer, substring(y.description, 6, 4));

create index idx_fiscalyear_dates  on #fiscalyear_dates  (fiscalyear);


--------------------------------------------------------------------------------
-- dropped or declined programs
--------------------------------------------------------------------------------
select *
into #dropped_declined 
from
(
select cr.constituentid, rp.id as recognitionprogramid
from constituentrecognition cr join recognitionprogram rp on rp.id = cr.recognitionprogramid
where cr.status = 'dropped' and rp.id in 
(
'E3E34E7C-0CFA-4D0E-9132-E905F9CEE09C',	-- Lux Libertas Society
'8AD2418D-7AA2-4522-9AC5-D704D0EE420D', -- Chancellor's Council
'97098F19-215E-47CE-9557-C6A81B10B3FD',	-- Chancellor's Clubs
'D3BFF196-CC45-4B9F-8A83-3165760CA174',	-- Chancellor's Clubs (1-5 Yr)
'24F82E62-954E-4875-B47D-AB12F0CD95E6',	-- Chancellor's Clubs (6-10 Yr)
'15819536-8A89-4D42-A96A-B8B4FB5699C6'	-- Chancellor's Clubs (Student)
)

union

select crdp.constituentid, rp.id as recognitionprogramid
from recognitionprogram rp join constituentrecognitiondeclinedprogram crdp on rp.id = crdp.recognitionprogramid 
where rp.id in 
(
'E3E34E7C-0CFA-4D0E-9132-E905F9CEE09C',	-- Lux Libertas Society
'8AD2418D-7AA2-4522-9AC5-D704D0EE420D', -- Chancellor's Council
'97098F19-215E-47CE-9557-C6A81B10B3FD',	-- Chancellor's Clubs
'D3BFF196-CC45-4B9F-8A83-3165760CA174',	-- Chancellor's Clubs (1-5 Yr)
'24F82E62-954E-4875-B47D-AB12F0CD95E6',	-- Chancellor's Clubs (6-10 Yr)
'15819536-8A89-4D42-A96A-B8B4FB5699C6'	-- Chancellor's Clubs (Student)
) 
) x;

create index idx_dropped_declined  on #dropped_declined  (constituentid, recognitionprogramid);

--------------------------------------------------------------------------------
-- program join dates
-- join date from constituent recognition
--------------------------------------------------------------------------------
select
cr.constituentid, 
rp.name as program,
cr.joindate,
case when rp.type = 'lifetime giving' then 9999 else convert(integer, substring(y.description, 6, 4)) end as fiscalyear

into 
#program_joindate_all

from 
constituentrecognition cr
join recognitionprogram rp on rp.id = cr.recognitionprogramid
left join constituentrecognitiondeclinedprogram d on cr.constituentid = d.constituentid and cr.recognitionprogramid = d.recognitionprogramid
join glfiscalperiod p on cr.joindate between p.startdate and p.enddate
join glfiscalyear y on y.id = p.glfiscalyearid

where 
rp.id in
(
'E3E34E7C-0CFA-4D0E-9132-E905F9CEE09C',	-- Lux Libertas Society
'8AD2418D-7AA2-4522-9AC5-D704D0EE420D', -- Chancellor's Council
'97098F19-215E-47CE-9557-C6A81B10B3FD',	-- Chancellor's Clubs
'D3BFF196-CC45-4B9F-8A83-3165760CA174',	-- Chancellor's Clubs (1-5 Yr)
'24F82E62-954E-4875-B47D-AB12F0CD95E6',	-- Chancellor's Clubs (6-10 Yr)
'15819536-8A89-4D42-A96A-B8B4FB5699C6'	-- Chancellor's Clubs (Student)
)
and d.constituentid is null
and cr.status <> 'dropped';


select constituentid, program, min(joindate) joindate, case when month(min(joindate)) < 7 then year(min(joindate)) else year(min(joindate)) + 1 end as fiscalyear
into #program_joindate
from #program_joindate_all
group by  constituentid, program, fiscalyear;

create index idx_program_joindate on #program_joindate (constituentid, program, fiscalyear);

--------------------------------------------------------------------------------
-- program revenue
-- giving information used to calculate if donors meet program criteria
--------------------------------------------------------------------------------
select *
into #program_revenue
from 
(
select
rr.constituentid, 
c.lookupid,
c.name,
c.keyname as lastname,
f.id as transactionid,
f.calculateduserdefinedid as revenue_id,
rr.effectivedate,
rrt.description as recognition_type,
convert(integer, substring(y.description, 6, 4)) as fiscalyear,
p.startdate as tx_fystartdate, 
p.enddate as tx_fyenddate,
l.id as lineitemid,
f.type as tx_type,
sx.type as revenue_type,
sx.application,
rr.amount,
c2.id as legal_constituentid,
c2.lookupid as legal_lookupid,
c2.name as legal_name

from
financialtransaction f
join financialtransactionlineitem l on l.financialtransactionid = f.id
join revenuesplit_ext sx on sx.id = l.id
join revenuerecognition rr on rr.revenuesplitid = l.id --and rr.revenuerecognitiontypecodeid in('17B72AB2-8FC8-47D7-A4B0-839D8D29B2FD', '643AF12A-4ED4-420F-AAB5-31F6A3525D1E') -- self, self third party
join revenuerecognitiontypecode rrt on rr.revenuerecognitiontypecodeid = rrt.id
join constituent c on c.id = rr.constituentid
join constituent c2 on c2.id = f.constituentid
join dbo.designation d on d.id = sx.designationid
join dbo.designationlevel dl on dl.id = d.designationlevel1id
join dbo.site s on s.id = dl.siteid
join glfiscalperiod p on rr.effectivedate between p.startdate and p.enddate
join glfiscalyear y on p.glfiscalyearid = y.id

where
f.deletedon is null
and l.deletedon is null
and l.typecode <> 1
and rr.amount > 0
and s.siteid <> 98
and f.date <= getdate()
and 
(
(
f.type in ('payment', 'matching gift claim')
and sx.type = 'gift'
and sx.application in ('recurring gift', 'donation', 'pledge', 'planned gift', 'matching gift', 'event registration')
)
or f.type in ('planned gift')
)
) x;

create index idx_program_revenue on #program_revenue (constituentid, fiscalyear, effectivedate, tx_type, lineitemid);

--------------------------------------------------------------------------------
-- degrees
-- education information used to calculate if donors meet program criteria
--------------------------------------------------------------------------------
select *
into #degree
from
(
select
eh.constituentid, min(eh.classof) classof, 'graduated' as type

from 
educationalhistory eh
join educationalhistorystatus es on eh.educationalhistorystatusid = es.id
join educationalprogramcode epc on eh.educationalprogramcodeid = epc.id

where
eh.educationalinstitutionid = '67673E25-CAF3-4D45-A4D2-EF633736B7D0'
and es.id in ('00000000-0000-0000-0000-000000000003', '00000000-0000-0000-0000-000000000004')  --incomplete, graduated
and epc.id = '13CAC54D-E659-423E-85E7-C12C0FFAD607' --bachelors degree

group by 
eh.constituentid

union

select
eh.constituentid, max(eh.classof) classof, 'attending' as type

from 
educationalhistory eh
join educationalhistorystatus es on eh.educationalhistorystatusid = es.id
join educationalprogramcode epc on eh.educationalprogramcodeid = epc.id

where
eh.educationalinstitutionid = '67673E25-CAF3-4D45-A4D2-EF633736B7D0'
and es.id = '00000000-0000-0000-0000-000000000002' -- currently attending
and epc.id = '13CAC54D-E659-423E-85E7-C12C0FFAD607' --bachelors degree

group by 
eh.constituentid
) x;

create index idx_degree on #degree (constituentid, classof, type);

--------------------------------------------------------------------------------
-- calculated program join dates from giving
-- exclude records where program was already calculated by davie
-- exclude dropped/declined records
--------------------------------------------------------------------------------

--------------------------------------------------------------------------------
-- lux
-- count all planned gifts whether paid or not.  other programs only count paid.
--------------------------------------------------------------------------------
select distinct t.constituentid, case when month(min(t.effectivedate)) < 7 then year(min(t.effectivedate)) else year(min(t.effectivedate)) + 1 end as fiscalyear, min(t.effectivedate) as date_achieved, 'Lux Libertas Society' as program
into #datelux
from 
(
select effectivedate, amount, constituentid,
	(
	select sum(amount) 
	from #program_revenue t2
	where t2.application <> 'planned gift' and t2.constituentid = t.constituentid and t2.effectivedate <= t.effectivedate
	) as cumulativeamount
from #program_revenue t where t.application <> 'planned gift'
) t 
left join #dropped_declined dd on t.constituentid = dd.constituentid and dd.recognitionprogramid = 'E3E34E7C-0CFA-4D0E-9132-E905F9CEE09C'
left join #program_joindate pj on t.constituentid = pj.constituentid and pj.program = 'Lux Libertas Society'

where 
cumulativeamount >= 1000000
and dd.constituentid is null
and pj.constituentid is null

group by
t.constituentid;


create index idx_datelux on #datelux (constituentid);

--------------------------------------------------------------------------------
-- chancellor's council
--------------------------------------------------------------------------------
select distinct constituentid, case when month(min(t.effectivedate)) < 7 then year(min(t.effectivedate)) else year(min(t.effectivedate)) + 1 end as fiscalyear, min(effectivedate) as date_achieved, 'Chancellor''s Council' as program
into #datecouncil
from (select effectivedate, amount, constituentid,
             (select sum(amount) from #program_revenue t2 where t2.tx_type <> 'planned gift' and t2.constituentid = t.constituentid and t2.effectivedate <= t.effectivedate) as cumulativeamount
      from #program_revenue t where t.tx_type <> 'planned gift'
     ) t
where cumulativeamount >= 100000
and constituentid not in 
(
select distinct cr.constituentid
from constituentrecognition cr join recognitionprogram rp on rp.id = cr.recognitionprogramid
where rp.name = 'Chancellor''s Council' and cr.status = 'dropped'
)
and constituentid not in 
(
select distinct crdp.constituentid
from recognitionprogram rp join constituentrecognitiondeclinedprogram crdp on rp.id = crdp.recognitionprogramid 
where rp.name = 'Chancellor''s Council'
)
and constituentid not in (select constituentid from #program_joindate where program in ('Chancellor''s Council', 'Lux Libertas Society'))
and constituentid not in (select constituentid from #datelux)
group by constituentid;

--------------------------------------------------------------------------------
-- chancellor's club
--------------------------------------------------------------------------------
select distinct t.constituentid, t.fiscalyear, min(effectivedate) as date_achieved, 'Chancellor''s Clubs' as program
into #dateclub
from 
(
select effectivedate, amount, constituentid, fiscalyear,
	(
	select sum(amount)
	from #program_revenue t2 
	where t2.tx_type <> 'planned gift' and t2.constituentid = t.constituentid and t2.fiscalyear = t.fiscalyear and t2.effectivedate <= t.effectivedate --and t2.tx_type = t.tx_type and t.tx_type in ('payment', 'matching gift claim')
	) as cumulativeamount

from #program_revenue t where t.tx_type <> 'planned gift'
) t
left join #program_joindate pj on t.constituentid = pj.constituentid and t.fiscalyear = pj.fiscalyear and pj.program = 'Chancellor''s Clubs'
left join #dropped_declined dd on t.constituentid = dd.constituentid and dd.recognitionprogramid = '97098F19-215E-47CE-9557-C6A81B10B3FD'

where 
cumulativeamount >= 2000
and dd.constituentid is null
and pj.constituentid is null

group by 
t.constituentid, t.fiscalyear

--------------------------------------------------------------------------------
-- chancellor's club (1-5)
--------------------------------------------------------------------------------
select distinct t.constituentid, t.fiscalyear, min(effectivedate) as date_achieved, 'Chancellor''s Clubs (1-5 Yr)' as program
into #dateclub1_5_x
from 
(
select effectivedate, amount, constituentid, fiscalyear,
	(
	select sum(amount) 
	from #program_revenue t2 
	where t2.tx_type <> 'planned gift' and t2.constituentid = t.constituentid and t2.fiscalyear = t.fiscalyear and t2.effectivedate <= t.effectivedate --and t2.tx_type = t.tx_type and t.tx_type in ('payment', 'matching gift claim')
	) as cumulativeamount

from 
#program_revenue t where t.tx_type <> 'planned gift'
) t
left join #program_joindate pj on t.constituentid = pj.constituentid and t.fiscalyear = pj.fiscalyear and pj.program = 'Chancellor''s Clubs (1-5 Yr)'
left join #dropped_declined dd on t.constituentid = dd.constituentid and dd.recognitionprogramid = 'D3BFF196-CC45-4B9F-8A83-3165760CA174'

where 
cumulativeamount >= 500
and dd.constituentid is null
and pj.constituentid is null

group by 
t.constituentid, t.fiscalyear;

create index idx_dateclub1_5_x on #dateclub1_5_x (constituentid, fiscalyear);


select distinct
c.constituentid, c.fiscalyear, c.date_achieved, d.classof, 'Chancellor''s Clubs (1-5 Yr)' as program

into 
#dateclub1_5

from
#dateclub1_5_x c
join #degree d on c.constituentid = d.constituentid 

where
c.fiscalyear - d.classof in (1,2,3,4,5);

--------------------------------------------------------------------------------
-- chancellor's club (6-10)
--------------------------------------------------------------------------------
select distinct t.constituentid, t.fiscalyear, min(effectivedate) as date_achieved, 'Chancellor''s Clubs (6-10 Yr)' as program
into #dateclub6_10_x
from 
(
select effectivedate, amount, constituentid, fiscalyear,
	(
	select sum(amount) 
	from #program_revenue t2 
	where t2.tx_type <> 'planned gift' and t2.constituentid = t.constituentid and t2.fiscalyear = t.fiscalyear and t2.effectivedate <= t.effectivedate --and t2.tx_type = t.tx_type and t.tx_type in ('payment', 'matching gift claim')
	) as cumulativeamount

from 
#program_revenue t where t.tx_type <> 'planned gift'
) t
left join #program_joindate pj on t.constituentid = pj.constituentid and t.fiscalyear = pj.fiscalyear and pj.program = 'Chancellor''s Clubs (6-10 Yr)'
left join #dropped_declined dd on t.constituentid = dd.constituentid and dd.recognitionprogramid = '24F82E62-954E-4875-B47D-AB12F0CD95E6'

where 
cumulativeamount >= 1000
and dd.constituentid is null
and pj.constituentid is null

group by 
t.constituentid, t.fiscalyear;

create index idx_dateclub6_10_x on #dateclub6_10_x (constituentid, fiscalyear);


select distinct
c.constituentid, c.fiscalyear, c.date_achieved, d.classof, 'Chancellor''s Clubs (6-10 Yr)' as program

into 
#dateclub6_10

from
#dateclub6_10_x c
join #degree d on c.constituentid = d.constituentid 

where
c.fiscalyear - d.classof in (6,7,8,9,10);


--------------------------------------------------------------------------------
-- chancellor's club (student)
--------------------------------------------------------------------------------
select distinct t.constituentid, t.fiscalyear as fiscalyear, min(effectivedate) as date_achieved, 'Chancellor''s Clubs (Student)' as program
into #dateclubstudent_x
from 
(
select effectivedate, amount, constituentid, fiscalyear,
	(
	select sum(amount) 
	from #program_revenue t2 
	where t2.tx_type <> 'planned gift' and t2.constituentid = t.constituentid and t2.fiscalyear = t.fiscalyear and t2.effectivedate <= t.effectivedate --and t2.tx_type = t.tx_type and t.tx_type in ('payment', 'matching gift claim')
	) as cumulativeamount

from 
#program_revenue t 
where t.tx_type <> 'planned gift' 
) t
left join #program_joindate pj on t.constituentid = pj.constituentid and t.fiscalyear = pj.fiscalyear and pj.program = 'Chancellor''s Clubs (Student)'
left join #dropped_declined dd on t.constituentid = dd.constituentid and dd.recognitionprogramid = '15819536-8A89-4D42-A96A-B8B4FB5699C6'

where
cumulativeamount >= 250
and dd.constituentid is null
and pj.constituentid is null

group by
t.constituentid, t.fiscalyear;

create index idx_dateclubstudent_x  on #dateclubstudent_x (constituentid, fiscalyear);


select distinct
c.constituentid, c.fiscalyear, c.date_achieved, d.classof, 'Chancellor''s Clubs (Student)' as program

into 
#dateclubstudent

from
#dateclubstudent_x c
join #degree d on c.constituentid = d.constituentid 

where
(c.fiscalyear = @CurrentFY and d.type = 'attending' and d.classof in (@CurrentFY, @CurrentFY+1, @CurrentFY+2, @CurrentFY+3))
or
(c.fiscalyear = @CurrentFY-1 and d.type = 'attending' and d.classof in (@CurrentFY-1, @CurrentFY-1+1, @CurrentFY-1+2, @CurrentFY-1+3));
--or
--(c.fiscalyear < @CurrentFY-1 and d.type = 'graduated');


--------------------------------------------------------------------------------
-- merged calculated and constituent recognition data
--------------------------------------------------------------------------------
select *
into #merged0
from
(
select constituentid, program, joindate, fiscalyear as joinyear, 'Constituent Recognition' as source, program as program_notation
from #program_joindate

union

select constituentid, program, date_achieved, case when month(date_achieved)  < 7 then year(date_achieved) else  year(date_achieved) + 1 end as joinyear, 'Calculated' as source, program + '*' as program_notation
from #datelux

union

select constituentid, program, date_achieved, case when month(date_achieved)  < 7 then year(date_achieved) else  year(date_achieved) + 1 end as joinyear,  'Calculated' as source, program + '*' as program_notation
from #datecouncil

union

select constituentid, program, date_achieved, fiscalyear as joinyear, 'Calculated' as source, program + '*' as program_notation
from #dateclub

union

select constituentid, program, date_achieved, fiscalyear as joinyear,  'Calculated' as source, program + '*' as program_notation
from #dateclub6_10

union

select constituentid, program, date_achieved, fiscalyear as joinyear,  'Calculated' as source, program + '*' as program_notation
from #dateclub1_5

union

select constituentid, program, date_achieved, fiscalyear as joinyear,  'Calculated' as source, program + '*' as program_notation
from #dateclubstudent
) xl


create index idx_merged0 on #merged0 (constituentid, program, joinyear, joindate);


select *
into #merged
from
(
select distinct
a.joinyear as yr, a.constituentid, a.program, a.joindate, a.joinyear, a.source, a.program_notation,
case when b.constituentid is null then 'New Member' else 'Prior Member' end as member_status,
fyd.startdate as fystartdate, fyd.enddate as fyenddate,
case
when month(a.joindate) = 7 then '01-Jul'
when month(a.joindate) = 8 then '02-Aug'
when month(a.joindate) = 9 then '03-Sep'
when month(a.joindate) = 10 then '04-Oct'
when month(a.joindate) = 11 then '05-Nov'
when month(a.joindate) = 12 then '06-Dec'
when month(a.joindate) = 1 then '07-Jan'
when month(a.joindate) = 2 then '08-Feb'
when month(a.joindate) = 3 then '09-Mar'
when month(a.joindate) = 4 then '10-Apr'
when month(a.joindate) = 5 then '11-May'
when month(a.joindate) = 6 then '12-Jun'
end as fiscaljoinmonth

from #merged0 a
left join #merged0 b on a.constituentid = b.constituentid and a.program = b.program and a.joinyear > b.joinyear
join #fiscalyear_dates fyd on a.joinyear = fyd.fiscalyear

where
a.program not in ('Lux Libertas Society', 'Chancellor''s Council')
--and a.joinyear between @CurrentFY-2 and @CurrentFY

union

select 
y.yr, m.*, case when y.yr = m.joinyear then 'New Member' when y.yr > m.joinyear then 'Prior Member' else '' end as member_status,
fyd.startdate as fystartdate, fyd.enddate as fyenddate,
case
when month(m.joindate) = 7 then '01-Jul'
when month(m.joindate) = 8 then '02-Aug'
when month(m.joindate) = 9 then '03-Sep'
when month(m.joindate) = 10 then '04-Oct'
when month(m.joindate) = 11 then '05-Nov'
when month(m.joindate) = 12 then '06-Dec'
when month(m.joindate) = 1 then '07-Jan'
when month(m.joindate) = 2 then '08-Feb'
when month(m.joindate) = 3 then '09-Mar'
when month(m.joindate) = 4 then '10-Apr'
when month(m.joindate) = 5 then '11-May'
when month(m.joindate) = 6 then '12-Jun'
end as fiscaljoinmonth

from 
#years y, #merged0 m, #fiscalyear_dates fyd 

where 
m.program in ('Lux Libertas Society', 'Chancellor''s Council') 
and y.yr >= m.joinyear
and m.joinyear = fyd.fiscalyear
) x;



--------------------------------------------------------------------------------
-- remove revenue records from temp table that don't relate to program recognition
--------------------------------------------------------------------------------
delete from #program_revenue where constituentid not in (select constituentid from #merged);


--------------------------------------------------------------------------------
-- join program information to revenue information
--------------------------------------------------------------------------------
select *
into #records
from
(
select distinct
m.yr,
pr.constituentid as recognition_id,
pr.lookupid as recognition_pid,
pr.name as recognition_name,
pr.lastname as recognition_lastname,
demo.householdlookupid as hh_lookupid,
case when demo.householdlookupid = '' then pr.lookupid else demo.householdlookupid end as hh_dedupeid,
pr.legal_constituentid,
pr.legal_lookupid,
pr.legal_name,
m.program,
m.program_notation,
m.joinyear,
m.fystartdate, 
m.fyenddate,
m.joindate,
m.fiscaljoinmonth,
m.member_status,
pr.recognition_type,
pr.effectivedate,
pr.revenue_id,
pr.lineitemid,
pr.amount,
pr.tx_type,
pr.revenue_type,
pr.application,
pr.fiscalyear as tx_fiscalyear,
pr.tx_fystartdate, 
pr.tx_fyenddate,
d.userid as designation_number,
d.vanityname as designation_name,
s.shortname as site,
upper(replace(replace(replace(dt.description, 'capital, ', ''), 'current, ', ''), 'endowment, ', '')) as fundtype,
upper(demo.primaryconstituency) primaryconstituency,
case when r.revenuesplitid is null then 'No' else 'Yes' end as campaign,
case when m.yr = pr.fiscalyear then 'Yes' else 'No' end as same_year,
'https://davie-crm.dev.unc.edu/bbappfx/webui/webshellpage.aspx?databasename=BBInfinity#pageType=p&pageId=387f861b-6c03-486c-9ff5-9cc5bb7a5275&recordId=' +  cast(pr.transactionid as varchar(36)) as tx_url,
'https://davie-crm.dev.unc.edu/bbappfx/webui/webshellpage.aspx?databasename=BBInfinity#pageType=p&pageId=88159265-2b7e-4c7b-82a2-119d01ecd40f&recordId=' +  cast(pr.constituentid as varchar(36)) as constituent_url

from
#program_revenue pr
join revenuesplit_ext sx on sx.id = pr.lineitemid
join designation d on d.id = sx.designationid
join designationlevel dl on dl.id = d.designationlevel1id
join designationleveltype dt on dt.id = dl.designationleveltypeid
join designationusecode duc on d.designationusecodeid = duc.id
join site s on s.id = dl.siteid
left join revenuesplitcampaign r on r.revenuesplitid = pr.lineitemid and r.campaignid = '5BC55E1B-495F-44A9-B3E7-888D8B62C1C2'
join bbinfinity_rpt_bbdw.bbdw.fact_usr_unc_demographic demo on pr.constituentid = demo.constituentsystemid
join #merged m on pr.constituentid = m.constituentid

where
m.yr = pr.fiscalyear
--and m.yr between @CurrentFY-2 and @CurrentFY

union

select distinct
m.yr,
m.constituentid as recognition_id,
demo.constituentlookupid as recognition_pid,
demo.constituentfullname as recognition_name,
demo.constituentlastname as recognition_lastname,
demo.householdlookupid as hh_lookupid,
case when demo.householdlookupid = '' then pr.lookupid else demo.householdlookupid end as hh_dedupeid,
null as legal_constituentid,
null as legal_lookupid,
null as legal_name,
m.program,
m.program_notation,
m.joinyear,
m.fystartdate, 
m.fyenddate,
m.joindate,
m.fiscaljoinmonth,
m.member_status,
null as recognition_type,
null as effectivedate,
null as revenue_id,
null as lineitemid,
null as amount,
null as tx_type,
null as revenue_type,
null as application,
null as tx_fiscalyear,
null as tx_fystartdate, 
null as tx_fyenddate,
null as designation_number,
null as designation_name,
null as site,
null as fundtype,
upper(demo.primaryconstituency) primaryconstituency,
null as campaign,
'Yes' as same_year,
null as tx_url,
'https://davie-crm.dev.unc.edu/bbappfx/webui/webshellpage.aspx?databasename=BBInfinity#pageType=p&pageId=88159265-2b7e-4c7b-82a2-119d01ecd40f&recordId=' +  cast(m.constituentid as varchar(36)) as constituent_url

from
#merged m
join bbinfinity_rpt_bbdw.bbdw.fact_usr_unc_demographic demo on m.constituentid = demo.constituentsystemid
left join #program_revenue pr on pr.constituentid = m.constituentid and pr.fiscalyear = m.yr

where
pr.constituentid is null
--and m.yr between @CurrentFY-2 and @CurrentFY
) x;


--------------------------------------------------------------------------------
-- remove council records after lux join date
--------------------------------------------------------------------------------
delete r
from #records r 
join (select recognition_pid, program, min(joinyear) min_lux_joinyear from #records where program = 'Lux Libertas Society'group by recognition_pid, program) l
on r.recognition_pid = l.recognition_pid and r.program = 'Chancellor''s Council' and r.yr >= l.min_lux_joinyear



select *, 
case when lineitemid is null then 0 
else row_number() over(partition by program, same_year, lineitemid order by hh_lookupid, recognition_type) end as seq
from #records




