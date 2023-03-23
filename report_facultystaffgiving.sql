


--------------------------------------------------------------
-- get current and prior fiscal year dates
--------------------------------------------------------------
declare @CurrentFY smallint;
declare @CurrentFYStart date;
declare @CurrentFYEnd date;
declare @CurrentActualFYEnd date;
declare @PriorFYStart date;
declare @PriorFYEnd date;
declare @PriorActualFYEnd date;
declare @OverrideDates bit;

set @OverrideDates = (select case when month(getdate()) = 7 and day(getdate()) < 11 then 1 else 0 end)
set @CurrentFY = (select y.yearid fiscalyear from dbo.glfiscalperiod p join glfiscalyear y on p.glfiscalyearid = y.id where getdate() between p.startdate and p.enddate);
set @CurrentFY = (select case when @OverrideDates = 1 then @CurrentFY - 1 else @CurrentFY end)
set @CurrentFYStart = (select min(p.startdate) from glfiscalyear y join glfiscalperiod p on y.id = p.glfiscalyearid where y.yearid = @CurrentFY)
set @CurrentFYEnd = getdate()
set @CurrentActualFYEnd = (select max(p.enddate) from glfiscalyear y join glfiscalperiod p on y.id = p.glfiscalyearid where y.yearid = @CurrentFY)
set @CurrentFYEnd = (select case when @OverrideDates = 1  then @CurrentActualFYEnd else @CurrentFYEnd end)
set @PriorFYStart = (select min(p.startdate) from glfiscalyear y join glfiscalperiod p on y.id = p.glfiscalyearid where y.yearid = @CurrentFY-1)
set @PriorFYEnd = (select dateadd(year, -1, @CurrentFYEnd))
set @PriorActualFYEnd = (select max(p.enddate) from glfiscalyear y join glfiscalperiod p on y.id = p.glfiscalyearid where y.yearid = @CurrentFY-1)
set @PriorFYEnd = (select case when @OverrideDates = 1  then @PriorActualFYEnd else @PriorFYEnd end)

--------------------------------------------------------------
-- drop tables (for testing)
--------------------------------------------------------------
--drop table #faculty_staff;
--drop table #gifts;
--drop table #detail;
--drop table #detail_cmpn;
--drop table #detail_cmpn2;
--drop table #detail_cmpn3;
--drop table #adjunct;

--------------------------------------------------------------
-- get adjunct records
--------------------------------------------------------------
select * into #adjunct from
(
select 
constituentid,
min(dateadded) as min_datefrom,
max(isnull(deleteddate, '9999-01-01')) as max_dateto,
case when min(enteronservicedate) > @PriorActualFYEnd or max(isnull(deleteddate, '9999-01-01')) < @PriorFYStart then 'No' else 'Yes' end as PriorFY_adjunct,	
case when min(enteronservicedate) > @CurrentActualFYEnd or max(isnull(deleteddate, '9999-01-01')) < @CurrentFYStart then 'No' else 'Yes' end as CurrentFY_adjunct

from 
usr_unc_employeehistory

where 
title like '%adjunct%'

group by
constituentid
)x;
--------------------------------------------------------------
-- get faculty/staff records
--------------------------------------------------------------
select * into #faculty_staff from
(
select 
ch.householdid,
c2.lookupid as hh_lookupid,
c.id as constituentid,
c.lookupid,
c.firstname + ' ' + c.keyname as name,
c.keyname as lastname,
cc.description as constituency_type,
isnull(cy.datefrom, cy.dateadded) as constituency_start_date,
isnull(cy.dateto, '9999-01-01') as constituency_end_date,
case when cy.datefrom > @PriorActualFYEnd or cy.dateto < @PriorFYStart then 'No' else 'Yes' end as PriorFY_facultystaff,	
case when cy.datefrom > @CurrentActualFYEnd or cy.dateto < @CurrentFYStart then 'No' else 'Yes' end as CurrentFY_facultystaff,
a.min_datefrom as adjunct_start_date,
a.max_dateto as adjunct_end_date,
isnull(a.PriorFY_adjunct, 'No') as PriorFY_adjunct,	
isnull(a.CurrentFY_adjunct, 'No') as CurrentFY_adjunct

from
constituent c
join constituency cy on c.id = cy.constituentid
join constituencycode cc on cy.constituencycodeid = cc.id
left join constituenthousehold ch on c.id = ch.id
left join constituent c2 on ch.householdid = c2.id
left join #adjunct a on c.id = a.constituentid

where
cc.description = 'Faculty/Staff' 
)x;

--------------------------------------------------------------
-- get gift records
--------------------------------------------------------------
select * into #gifts from
(
select distinct
f.id as txid,
c.id as constituentid,
rr.effectivedate as gift_date,
case 
when rr.effectivedate between @PriorFYStart and @PriorFYEnd then 'PriorFYGift' else'CurFYGift' 
end as FY,
rr.amount as gift_amt,
f.calculateduserdefinedid as rev_id,
s.shortname as site,
d.userid as designationid,
d.vanityname as designation_name,
a.name as appeal,
am.name as mailing,
re.sourcecode,
cmpn.name as campaign,
f.type as txtype,
sx.application,
re.receipttype,
case 
when f.type = 'payment' and sx.application = 'recurring gift' then 1
when re.receipttype = 'consolidated' then 1
else 0 
end as recurring,
dt.description as purposelevel,
'https://davie-crm.dev.unc.edu/bbappfx/webui/webshellpage.aspx?databasename=BBInfinity#pageType=p&amp;pageId=387f861b-6c03-486c-9ff5-9cc5bb7a5275&amp;recordId=' + cast(f.id as varchar(36)) as url

from 
dbo.revenuerecognition rr 
join dbo.revenuerecognitiontypecode rcd on rcd.id = rr.revenuerecognitiontypecodeid --and rcd.description like 'self%'
join dbo.financialtransactionlineitem l on rr.revenuesplitid = l.id
join dbo.financialtransaction f on l.financialtransactionid = f.id
join dbo.revenuesplit_ext sx on sx.id = l.id
join dbo.designation d on d.id = sx.designationid
join dbo.designationlevel dl on dl.id = d.designationlevel1id
join dbo.designationleveltype dt on dt.id = dl.designationleveltypeid
join dbo.site s on s.id = dl.siteid
join dbo.constituent c on c.id = rr.constituentid
left join revenue_ext re on f.id = re.id
left join appeal a on re.appealid = a.id
left join appealmailing am on a.id = re.mailingid
left join dbo.revenuesplitcampaign rc on rc.revenuesplitid = l.id
left join dbo.campaign cmpn on cmpn.id = rc.campaignid

where
f.type = 'Payment'
and sx.type = 'Gift'
and ((rr.effectivedate between @CurrentFYStart and @CurrentFYEnd) or (rr.effectivedate between @PriorFYStart and @PriorFYEnd))
)x;

select * into #detail from
(
select distinct
fs.householdid as HouseholdID,
fs.hh_lookupid as HouseholdLookupID,
fs.constituentid as ConstituentID,
fs.lookupid as PID, 
fs.name as Name,
fs.lastname as LastName,
fs.constituency_start_date as FacultyStaff_StartDate,
fs.constituency_end_date as FacultyStaff_EndDate,
fs.PriorFY_facultystaff,
fs.CurrentFY_facultystaff,
fs.adjunct_start_date as Adjunct_StartDate,
fs.adjunct_end_date as Adjunct_EndDate,
fs.PriorFY_adjunct,
fs.CurrentFY_adjunct,
g.txid as TxID,
g.rev_id as RevID,
g.gift_date as GiftDate,
g.FY,
case when g.gift_date between fs.constituency_start_date and fs.constituency_end_date then 'Yes' else 'No' end as [Gift During Constituency],
case when g.gift_date between fs.adjunct_start_date and fs.adjunct_end_date then 'Yes' else 'No' end as [Gift During Adjunct],
g.gift_amt as GiftAmount,
g.site as Site,
g.designationid as DesignationID,
g.designation_name as DesignationName,
g.appeal as AppealName,
g.mailing as MailingName,
g.sourcecode as SourceCode,
g.campaign,
g.txtype, 
g.application,
g.receipttype,
g.recurring,
g.purposelevel as PurposeLevel,
g.url

from 
#faculty_staff fs left join #gifts g on fs.constituentid = g.constituentid
)x;

--------------------------------------------------------------
-- add group concat campaign column
--------------------------------------------------------------
select * into #detail_cmpn from
(
select distinct
d1.HouseholdID,
d1.HouseholdLookupID,
d1.ConstituentID,
d1.PID,
d1.Name,
d1.LastName,
d1.FacultyStaff_StartDate,
d1.FacultyStaff_EndDate,
d1.PriorFY_facultystaff,
d1.CurrentFY_facultystaff,
d1.Adjunct_StartDate,
d1.Adjunct_EndDate,
d1.PriorFY_adjunct,
d1.CurrentFY_adjunct,
d1.TxID,
d1.RevID,
d1.GiftDate,
d1.FY,
d1.[Gift During Constituency],
d1.[Gift During Adjunct],
d1.GiftAmount,
d1.Site,
d1.DesignationID,
d1.DesignationName,
d1.AppealName,
d1.MailingName,
d1.SourceCode,
d1.PurposeLevel,
d1.txtype,
d1.application,
d1.receipttype,
d1.recurring,
d1.url,

stuff
((
select distinct cast(',' as varchar(max)) + rtrim(ltrim(d2.campaign))
from #detail d2
where d1.revid = d2.revid
for xml path('')
), 1, 1, '') as Campaign

from 
#detail d1
)x;

--------------------------------------------------------------
-- flag one transaction per household to count in totals
--------------------------------------------------------------
select * into #detail_cmpn2 from
(
select row_number() over(partition by householdid, txid, designationid, giftamount order by [Gift During Constituency] desc, [Gift During Adjunct] desc) as HouseholdRow, *
from #detail_cmpn
)x;


select * into #detail_cmpn3 from
(
select 
case when householdid is null then 1 else HouseholdRow end as HouseholdRowCnt, *
from #detail_cmpn2
)x;

select * from #detail_cmpn3