

declare @email nvarchar(100) = '*****@unc.edu'
declare @cc nvarchar(100) = '*****@unc.edu'
declare @baseurl nvarchar(100) = dbo.UFN_USR_UNC_BASEURLLINK();
declare @currYear smallint = datepart(yy,getdate());
declare @fyStart date = dbo.ufn_date_thisfiscalyear_Firstday(getdate(),0);
declare @fyEnd date = dbo.ufn_date_thisfiscalyear_Lastday(getdate(),0);

with records as
(
select distinct 
c.lookupid cid,
@baseurl + N'/webui/webshellpage.aspx?databasename=BBInfinity#pageType=p&amp;pageId=88159265-2b7e-4c7b-82a2-119d01ecd40f&amp;recordId=' + convert(nvarchar(36),C.ID) AS CURL,
c.NAME cnm,
rp.NAME,
cr.status,
rev.amt,
m.maxjoindate,
cr.dateadded,
cr.datechanged,
ca.username as add_user,
ca2.username as change_user,
@baseurl + N'/webui/webshellpage.aspx?databasename=BBInfinity#pageType=p&amp;pageId=0bdfb03b-f3e8-4056-b64a-4b2210603484&amp;recordId=' + convert(nvarchar(36),cr.ID) AS rURL

from 
dbo.constituent c
join dbo.constituentrecognition cr on cr.constituentid = c.id
join dbo.recognitionprogram rp on rp.id = cr.recognitionprogramid
join dbo.changeagent ca on cr.addedbyid = ca.id
left join dbo.changeagent ca2 on cr.changedbyid = ca2.id
left join (select distinct constituentid from constituency where constituencycodeid = '15B40BCF-7092-44F7-9E47-07E5DD780610') parent on c.id = parent.constituentid -- parent-current student
join 
(select constituentid, recognitionprogramid, max(joindate) as maxjoindate
from constituentrecognition
where joindate between @fystart and @fyend
group by constituentid, recognitionprogramid) m 
on cr.constituentid = m.constituentid and cr.recognitionprogramid = m.recognitionprogramid and cr.joindate = m.maxjoindate

outer apply (select sum(rr.amount) amt
from dbo.revenuerecognition rr 
join dbo.revenuesplit_ext sx on sx.id = rr.revenuesplitid
join dbo.designation d on d.id = sx.designationid
join dbo.designationlevel dl on dl.id = d.designationlevel1id
join dbo.site s on s.id = dl.siteid
where rr.constituentid = c.id
and rr.effectivedate between @fystart and @fyend
) rev

where 
rp.name = 'Carolina Parents Leadership Society'
and cr.joindate between @fystart and @fyend
and (isnull(rev.amt,0) < 5000 or parent.constituentid is null)
and ca.username = 'AD\adv_BBWS' 
and ca2.username = 'AD\adv_BBWS'
)

select @email email, @cc cc
where 1 <= (select count(*) from records)