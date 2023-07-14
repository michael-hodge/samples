

--------------------------------------------------------------------------------------
-- testing
--------------------------------------------------------------------------------------
--drop table #ladder_definition;
--drop table #ladder_giving;
--drop table #ladder_last_giving_year;
--drop table #ladder_giving_all;
--drop table #ladder_lag;
--drop table #ladder_group;
--drop table #ladder_calc_lastyeargiving;
--drop table #ladder_calc_totalgiving;
--drop table #ladder_calc_median;
--drop table #ladder_calc_currentyeargiving;
--drop table #ladder_calc_basis;
--drop table #ladder_group;
--drop table #ladder_test;

--declare @changeagentid uniqueidentifier = null;
--declare @currentdate datetime = getdate();

--if @changeagentid is null
--exec dbo.USP_CHANGEAGENT_GETORCREATECHANGEAGENT @CHANGEAGENTID output;
--------------------------------------------------------------------------------------


truncate table usr_unc_askladder;

---------------------------------------------------------------
-- get current and prior 10 fiscal years
---------------------------------------------------------------
declare @currentfyid smallint;
declare @fy0 smallint;
declare @fy1 smallint;
declare @fy2 smallint;
declare @fy3 smallint;
declare @fy4 smallint;
declare @fy5 smallint;
declare @fy6 smallint;
declare @fy7 smallint;
declare @fy8 smallint;
declare @fy9 smallint;
declare @fy10 smallint;
declare @fy11 smallint;

set @currentfyid = 
(
select y.yearid fiscalyear 
from glfiscalperiod p join glfiscalyear y on p.glfiscalyearid = y.id 
where getdate() between p.startdate and p.enddate
);

set @fy0 = 2000 + @currentfyid;
set @fy1 = 2000 + @currentfyid-1;
set @fy2 = 2000 + @currentfyid-2;
set @fy3 = 2000 + @currentfyid-3;
set @fy4 = 2000 + @currentfyid-4;
set @fy5 = 2000 + @currentfyid-5;
set @fy6 = 2000 + @currentfyid-6;
set @fy7 = 2000 + @currentfyid-7;
set @fy8 = 2000 + @currentfyid-8;
set @fy9 = 2000 + @currentfyid-9;
set @fy10 = 2000 + @currentfyid-10;
set @fy11 = 2000 + @currentfyid-11;


---------------------------------------------------------------
-- create ask ladder definitions
-- ask types:  
--   static = value.  
--   add = add to value.  
--   multiply = multiply by value.  
--   addR = add to value and round to nearest 5.  
--   multiplyR = multiply by value and round to nearest 5.
---------------------------------------------------------------
create table #ladder_definition
(
ladder_group varchar(100),
minrange money, 
maxrange money, 
ask1 numeric(12,2), 
ask2 numeric(12,2), 
ask3 numeric(12,2),
ask1type varchar(10), 
ask2type varchar(10), 
ask3type varchar(10)
);


insert into #ladder_definition
values 
('MultiYearRenewal',			1,		24.99,		75,		50,		35,		'static',		'static',		'static'),
('MultiYearRenewal',			25,		99.99,		2,		1.25,	1,		'multiplyR',	'multiplyR',	'multiplyR'),
('MultiYearRenewal',			100,	249.99,		2,		1.15,	1,		'multiplyR',	'multiplyR',	'multiplyR'),
('MultiYearRenewal',			250,	749.99,		1.8,	1.15,	1,		'multiplyR',	'multiplyR',	'multiplyR'),
('MultiYearRenewal',			750,	999.99,		1.8,	1000,	1,		'multiplyR',	'multiplyR',	'multiplyR'),
('MultiYearRenewal',			1000,	1699.99,	2000,	1.1,	1,		'static',		'multiplyR',	'multiplyR'),
('MultiYearRenewal',			1700,	1999.99,	5000,	2000,	1,		'static',		'static',		'multiplyR'),
('MultiYearRenewal',			2000,	4999.99,	10000,	5000,	2000,	'static',		'static',		'static'),
('MultiYearRenewal',			5000,	9999.99,	15000,	10000,	5000,	'static',		'static',		'static'),
('MultiYearRenewal',			10000,	24999.99,	25000,	15000,	1000,	'static',		'static',		'static'),
('MultiYearRenewal',			25000,	9999999999,	0,		0,		0,		'static',		'static',		'static'),

('New',							1,		24.99,		50,		350,	25,		'static',		'static',		'static'),
('New',							25,		49.99,		2,		1.25,	1,		'multiplyR',	'multiplyR',	'multiplyR'),
('New',							50,		99.99,		1.5,	1.25,	1,		'multiplyR',	'multiplyR',	'multiplyR'),
('New',							100,	249.99,		50,		25,		1,		'addR',			'addR',			'multiplyR'),
('New',							250,	499.99,		100,	50,		1,		'addR',			'addR',			'multiplyR'),
('New',							500,	749.99,		100,	50,		1,		'addR',			'addR',			'multiplyR'),
('New',							750,	999.99,		1250,	1000,	1,		'static',		'static',		'multiply'),
('New',							1000,	1499.99,	2000,	1500,	1,		'static',		'static',		'multiply'),
('New',							1500,	1999.99,	2500,	2000,	1,		'static',		'static',		'multiply'),
('New',							2000,	4999.99,	7500,	5000,	2000,	'static',		'static',		'static'),
('New',							5000,	9999.99,	10000,	7500,	5000,	'static',		'static',		'static'),
('New',							10000,	24999.99,	15000,	12500,	10000,	'static',		'static',		'static'),
('New',							25000,	9999999999,	0,		0,		0,		'static',		'static',		'static'),

('Reactivate',					1,		24.99,		50,		350,	25,		'static',		'static',		'static'),
('Reactivate',					25,		49.99,		2,		1.25,	1,		'multiplyR',	'multiplyR',	'multiplyR'),
('Reactivate',					50,		99.99,		1.5,	1.25,	1,		'multiplyR',	'multiplyR',	'multiplyR'),
('Reactivate',					100,	249.99,		50,		25,		1,		'addR',			'addR',			'multiplyR'),
('Reactivate',					250,	499.99,		100,	50,		1,		'addR',			'addR',			'multiplyR'),
('Reactivate',					500,	749.99,		100,	50,		1,		'addR',			'addR',			'multiplyR'),
('Reactivate',					750,	999.99,		1250,	1000,	1,		'static',		'static',		'multiply'),
('Reactivate',					1000,	1499.99,	2000,	1500,	1,		'static',		'static',		'multiply'),
('Reactivate',					1500,	1999.99,	2500,	2000,	1,		'static',		'static',		'multiply'),
('Reactivate',					2000,	4999.99,	7500,	5000,	2000,	'static',		'static',		'static'),
('Reactivate',					5000,	9999.99,	10000,	7500,	5000,	'static',		'static',		'static'),
('Reactivate',					10000,	24999.99,	15000,	12500,	10000,	'static',		'static',		'static'),
('Reactivate',					25000,	9999999999,	0,		0,		0,		'static',		'static',		'static'),

('1YearLapse',					1,		24.99,		50,		350,	25,		'static',		'static',		'static'),
('1YearLapse',					25,		49.99,		2,		1.25,	1,		'multiplyR',	'multiplyR',	'multiplyR'),
('1YearLapse',					50,		99.99,		1.5,	1.25,	1,		'multiplyR',	'multiplyR',	'multiplyR'),
('1YearLapse',					100,	249.99,		50,		25,		1,		'addR',			'addR',			'multiplyR'),
('1YearLapse',					250,	499.99,		100,	50,		1,		'addR',			'addR',			'multiplyR'),
('1YearLapse',					500,	749.99,		100,	50,		1,		'addR',			'addR',			'multiplyR'),
('1YearLapse',					750,	999.99,		1250,	1000,	1,		'static',		'static',		'multiply'),
('1YearLapse',					1000,	1499.99,	2000,	1500,	1,		'static',		'static',		'multiply'),
('1YearLapse',					1500,	1999.99,	2500,	2000,	1,		'static',		'static',		'multiply'),
('1YearLapse',					2000,	4999.99,	7500,	5000,	2000,	'static',		'static',		'static'),
('1YearLapse',					5000,	9999.99,	10000,	7500,	5000,	'static',		'static',		'static'),
('1YearLapse',					10000,	24999.99,	15000,	12500,	10000,	'static',		'static',		'static'),
('1YearLapse',					25000,	9999999999,	0,		0,		0,		'static',		'static',		'static'),

('ShortLapse',					1,		49.99,		50,		35,		25,		'static',		'static',		'static'),
('ShortLapse',					50,		1899.99,	1,		.5,		.25,	'multiplyR',	'multiplyR',	'multiplyR'),
('ShortLapse',					1900,	2499.99,	2000,	1500,	1000,	'static',		'static',		'static'),
('ShortLapse',					2500,	4999.99,	2500,	2000,	1500,	'static',		'static',		'static'),
('ShortLapse',					5000,	9999.99,	5000,	2500,	2000,	'static',		'static',		'static'),
('ShortLapse',					10000,	24999.99,	10000,	5000,	2500,	'static',		'static',		'static'),
('ShortLapse',					25000,	9999999999,	0,		0,		0,		'static',		'static',		'static'),

('Acquisition',					1,		24999.99,	100,	50,		25,		'static',		'static',		'static'),
('Acquisition',					25000,	9999999999,	0,		0,		0,		'static',		'static',		'static'),

('LongLapse',					1,		24999.99,	100,	50,		25,		'static',		'static',		'static'),
('LongLapse',					25000,	9999999999,	0,		0,		0,		'static',		'static',		'static'),

('SecondAsk',					1,		99,			50,		25,		15,		'static',		'static',		'static'),
('SecondAsk',					100,	24999.99,	100,	50,		25,		'static',		'static',		'static'),
('SecondAsk',					25000,	9999999999,	100,	50,		25,		'static',		'static',		'static');

create index idx_ladder_definition on #ladder_definition (ladder_group, minrange, maxrange, ask1type, ask2type, ask3type);


---------------------------------------------------------------
-- get fiscal year amounts by constituent and site
-- from CAF FYXX Individuals and Spouses Donor Recognition Selection
---------------------------------------------------------------
select  	
r.constituentid,
r.siteid,
r.sitename,
case when month(r.effectivedate) < 7 then year(r.effectivedate) else year(r.effectivedate) + 1 end as fiscalyear,
sum(amount) amount

into
#ladder_giving

from 
v_query_constituent as c
left join usr_unc_v_query_constituentsimplerecognition as r on c.id = r.constituentid

where 
c.isindividual = 1
and c.deceased = 0
and c.isinactive = 0
and r.type = N'Gift'
and r.transactiontype = N'Payment'
and ((r.designationid not  in (N'6ab85ca2-b947-4941-b714-0de429be4296', N'bf833952-c953-4eda-8d9a-c1d2c6b3e3b7', N'0d127428-c702-45ff-aab3-6cf37755a6e0', N'70409353-87ab-44da-87f9-5da31b1c4564', N'44f35b91-3633-4d4d-9d9e-90856ef5e0a5', N'bfd9729e-2be0-4980-9ee0-3d67d8074118', N'1c95eea2-d58b-4ec1-913c-e4bc7909daa9')) or r.DESIGNATIONID is null)
and ((r.siteid not  in (N'c385f431-b701-4ead-9fd4-5821849d5005', N'420d550e-0670-4862-8d4a-2983fd1aae04', N'16e13730-0e0a-4ed9-8a81-bd8cb19dc6e6', N'd8b51327-758e-4e3a-ac9a-a1948766a924')) or r.SITEID is null)
and r.is_memorial_designation = 0
and ((r.paymentmethod <> N'Property') or r.paymentmethod is null or r.paymentmethod = '')
and ((r.giftinkindsubtypecodeid <> N'228dcdc3-71ea-4400-a307-c93b4e50e4ee') or r.giftinkindsubtypecodeid is null)

group by
r.constituentid,
r.siteid,
r.sitename,
case when month(r.effectivedate) < 7 then year(r.effectivedate) else year(r.effectivedate) + 1 end

having 
sum(amount) > 0

---------------------------------------------------------------
-- add row for overall giving
---------------------------------------------------------------
select * 
into #ladder_giving_all
from
(
select
convert(varchar(36), constituentid) + '_' +  '2704CFD1-6BB1-414D-8C1B-5A8BEE701946' as keyfield,
constituentid,
'2704CFD1-6BB1-414D-8C1B-5A8BEE701946' as siteid, 
'Overall' as sitename, 
fiscalyear,
sum(amount) as amount

from
#ladder_giving

group by
constituentid, fiscalyear

union

select convert(varchar(36), constituentid) + '_' +  convert(varchar(36), siteid) as keyfield, * from #ladder_giving
) x;

create index idx_ladder_giving_all on #ladder_giving_all (keyfield, constituentid, siteid, fiscalyear);

---------------------------------------------------------------
-- get last giving year
---------------------------------------------------------------
select constituentid, siteid, max(fiscalyear) as last_giving_year
into #ladder_last_giving_year
from #ladder_giving_all
group by constituentid, siteid;

create index idx_ladder_last_giving_year on #ladder_last_giving_year (constituentid, siteid);

---------------------------------------------------------------
-- calculate lag
---------------------------------------------------------------
--select keyfield, constituentid, siteid, fiscalyear, fiscalyear - lag(fiscalyear, 1) over (partition by constituentid, siteid order by constituentid, siteid, fiscalyear) as lag
--into #ladder_lag
--from #ladder_giving_all

--create index idx_ladder_lag on #ladder_lag (constituentid, siteid, fiscalyear, lag);


---------------------------------------------------------------
-- calculate ladder groups
---------------------------------------------------------------
create table #ladder_group
(
keyfield varchar(73),
constituentid uniqueidentifier,
siteid uniqueidentifier, 
ladder_group varchar(50) 
);

create index idx_ladder_group on #ladder_group (constituentid, siteid, ladder_group);


-- second ask
insert into #ladder_group
select keyfield, constituentid, siteid, 'SecondAsk' as ladder_group
from #ladder_giving_all
where fiscalyear = @fy0;


-- multi year renewals:  gift in last two years
insert into #ladder_group
select l.keyfield, l.constituentid, l.siteid, 'MultiYearRenewal' as ladder_group
from #ladder_giving_all l left join #ladder_group g on l.keyfield = g.keyfield
where fiscalyear in (@fy1, @fy2) and g.keyfield is null
group by l.keyfield, l.constituentid, l.siteid
having count(*) = 2;


-- new:  first gift was last year
insert into #ladder_group
select l.keyfield, l.constituentid, l.siteid, 'New' as ladder_group
from #ladder_giving_all l left join #ladder_group g on l.keyfield = g.keyfield
where g.keyfield is null
group by l.keyfield, l.constituentid, l.siteid
having min(l.fiscalyear) = @fy1;


-- 1 year lapse:  gift year before last, but none last year
insert into #ladder_group
select 
l.keyfield, l.constituentid, l.siteid, '1YearLapse' as ladder_group

from 
#ladder_giving_all l 
left join #ladder_giving_all l2 on l.keyfield = l2.keyfield and l2.fiscalyear = @fy1
left join #ladder_group g on l.keyfield = g.keyfield

where
l.fiscalyear = @fy2
and l2.keyfield is null
and g.keyfield is null;


-- reactivate:  gift last year, but not the year before
insert into #ladder_group
select 
l.keyfield, l.constituentid, l.siteid, 'Reactivate' as ladder_group

from 
#ladder_giving_all l 
left join #ladder_giving_all l2 on l.keyfield = l2.keyfield and l2.fiscalyear = @fy2
left join #ladder_group g on l.keyfield = g.keyfield

where
l.fiscalyear = @fy1
and l2.keyfield is null
and g.keyfield is null;


-- short lapse
insert into #ladder_group
select l.keyfield, l.constituentid, l.siteid, 'ShortLapse' as ladder_group
from #ladder_giving_all l left join #ladder_group g on l.keyfield = g.keyfield
where g.keyfield is null
group by l.keyfield, l.constituentid, l.siteid
having max(l.fiscalyear) in (@fy3, @fy4, @fy5, @fy6);


-- long lapse
insert into #ladder_group
select l.keyfield, l.constituentid, l.siteid, 'LongLapse' as ladder_group
from #ladder_giving_all l left join #ladder_group g on l.keyfield = g.keyfield
where g.keyfield is null
group by l.keyfield, l.constituentid, l.siteid
having max(l.fiscalyear) in (@fy7, @fy8, @fy9, @fy10, @fy11);


-- acquisition
insert into #ladder_group
select l.keyfield, l.constituentid, l.siteid, 'Acquisition' as ladder_group
from #ladder_giving_all l left join #ladder_group g on l.keyfield = g.keyfield
where g.keyfield is null
group by l.keyfield, l.constituentid, l.siteid
having max(l.fiscalyear) < @fy11;


---------------------------------------------------------------
-- get calculation basis values
---------------------------------------------------------------

-- last year giving
select constituentid, siteid, amount
into #ladder_calc_lastyeargiving
from #ladder_giving_all
where fiscalyear = @fy1

create index idx_ladder_calc_lastyeargiving on #ladder_calc_lastyeargiving (constituentid, siteid);

-- total giving
select constituentid, siteid, sum(amount) as amount
into #ladder_calc_totalgiving
from #ladder_giving_all
group by constituentid, siteid;

create index idx_ladder_calc_totalgiving on #ladder_calc_totalgiving (constituentid, siteid);

-- 5 year median giving
select distinct constituentid, siteid, percentile_cont(0.5) within group (order by amount) over (partition by constituentid, siteid) as amount
into #ladder_calc_median
from #ladder_giving_all
where fiscalyear in (@fy2, @fy3, @fy4, @fy5, @fy6)

create index idx_ladder_calc_median on #ladder_calc_median (constituentid, siteid);

-- current year giving
select constituentid, siteid, amount
into #ladder_calc_currentyeargiving
from #ladder_giving_all
where fiscalyear = @fy0

create index idx_ladder_calc_currentyeargiving on #ladder_calc_currentyeargiving (constituentid, siteid);

---------------------------------------------------------------
-- determine calculation basis
---------------------------------------------------------------
select 
g.*, isnull(m.amount, 0) as giving_5yr_median, isnull(l.amount, 0)  as giving_lastyr, isnull(c.amount, 0) as giving_currentyr, isnull(t.amount, 0)  as giving_total,
case 
when g.ladder_group = 'SecondAsk' then 'CurrentYearGiving'
when g.ladder_group = 'MultiYearRenewal' then 'LastYearGiving'
when g.ladder_group = 'New' then 'LastYearGiving'
when g.ladder_group = 'Reactivate' then 'LastYearGiving'
when g.ladder_group = '1YearLapse' then 'LastYearGiving'
when g.ladder_group = 'ShortLapse' then 'Last5YearsMedianGiving'
when g.ladder_group = 'Acquisition' then 'TotalGiving'
when g.ladder_group = 'LongLapse' then 'TotalGiving'
end as calcbasis,
case 
when g.ladder_group = 'SecondAsk' then c.amount
when g.ladder_group = 'MultiYearRenewal' then l.amount
when g.ladder_group = 'New' then l.amount
when g.ladder_group = 'Reactivate' then l.amount
when g.ladder_group = '1YearLapse' then l.amount
when g.ladder_group = 'ShortLapse' then m.amount
when g.ladder_group = 'Acquisition' then t.amount
when g.ladder_group = 'LongLapse' then t.amount
end as calcvalue

into
#ladder_calc_basis

from 
#ladder_group g
left join #ladder_calc_median m on g.constituentid = m.constituentid and g.siteid = m.siteid 
left join #ladder_calc_lastyeargiving l on g.constituentid = l.constituentid and g.siteid = l.siteid 
left join #ladder_calc_totalgiving t on g.constituentid = t.constituentid and g.siteid = t.siteid 
left join #ladder_calc_currentyeargiving c on g.constituentid = c.constituentid and g.siteid = c.siteid 

create index idx_ladder_calc_basis on #ladder_calc_basis (constituentid, siteid, ladder_group, calcbasis);


---------------------------------------------------------------
-- determine ask amounts
---------------------------------------------------------------
insert into usr_unc_askladder
(id, constituentid, siteid, ladder_group, giving_5yr_median, giving_lastyr, giving_currentyr, giving_total, calcbasis, calcvalue, minrange,
maxrange, ask1, ask2, ask3, ask1type, ask2type, ask3type, ask1amount, ask2amount, ask3amount, lastgiftyear, addedbyid, changedbyid, dateadded, datechanged)

select 
newid() as id, c.constituentid, c.siteid, c.ladder_group, c.giving_5yr_median, c.giving_lastyr, c.giving_currentyr, c.giving_total, 
c.calcbasis, c.calcvalue, d.minrange, d.maxrange, d.ask1, d.ask2, d.ask3, d.ask1type, d.ask2type, d.ask3type,
case
when d.ask1type = 'static' then ask1
when d.ask1type = 'multiplyR' then round((ask1 * c.calcvalue)/5,0) * 5
when d.ask1type = 'multiply' then ask1 * c.calcvalue
when d.ask1type = 'addR' then round((ask1 + c.calcvalue)/5,0) * 5
when d.ask1type = 'add' then ask1 + c.calcvalue
end as ask1amount,

case
when d.ask2type = 'static' then ask2
when d.ask2type = 'multiplyR' then round((ask2 * c.calcvalue)/5,0) * 5
when d.ask2type = 'multiply' then ask2 * c.calcvalue
when d.ask2type = 'addR' then round((ask2 + c.calcvalue)/5,0) * 5
when d.ask2type = 'add' then ask2 + c.calcvalue
end as ask2amount,

case
when d.ask3type = 'static' then ask3
when d.ask3type = 'multiplyR' then round((ask3 * c.calcvalue)/5,0) * 5
when d.ask3type = 'multiply' then ask3 * c.calcvalue
when d.ask3type = 'addR' then round((ask3 + c.calcvalue)/5,0) * 5
when d.ask3type = 'add' then ask3 + c.calcvalue
end as ask3amount,

isnull(l.last_giving_year, 0),

@changeagentid as addedbyid,
@changeagentid as changedbyid,
@currentdate as dateadded,
@currentdate as datechanged

from 
#ladder_calc_basis c 
join #ladder_definition d on c.ladder_group = d.ladder_group and c.calcvalue between d.minrange and d.maxrange
left join #ladder_last_giving_year l on c.constituentid = l.constituentid and c.siteid = l.siteid

where
c.calcvalue >= 1 and c.calcvalue < 25000
