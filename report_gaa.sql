
			
create procedure [dbo].[USR_UNC_USP_REPORT_GAA_CATALOG]
@pids nvarchar(max)
as
 begin



create table #pids_gaa_catalog (val nvarchar(100), sortorder int identity(1,1))
insert into #pids_gaa_catalog (val) 
select val from bbinfinity.dbo.USR_UNC_UFN_STRING_TO_TABLE (@pids,',',1);

-- involvement
with involvement as
(
select distinct a.constituentdimid,

stuff
((
select distinct cast('$' as varchar(max)) + rtrim(ltrim(educationalinvolvementname))
from bbinfinity_rpt_bbdw.bbdw.dim_educationalinvolvement b
where a.constituentdimid = b.constituentdimid  
for xml path('')
), 1, 1, '') as involvement

from bbinfinity_rpt_bbdw.bbdw.dim_educationalinvolvement a 
join bbinfinity_rpt_bbdw.bbdw.dim_constituent c on a.constituentdimid = c.constituentdimid
join #pids_gaa_catalog p on c.constituentlookupid = p.val
),

-- degree
degree as
(
select distinct a.constituentdimid,

stuff
((
select distinct cast('$' as varchar(max)) + rtrim(ltrim(case when b.educationdegree = 'No Educational Degree' then 'Degree Unknown' else b.educationdegree end
+ ' (' + case when b.classof = 0 then '' else cast(b.classof as varchar) end 
+ ' ' + case when b.educationinstitution = 'generic high school' then x.highschoolname else b.educationinstitution end + ')'))

from bbinfinity_rpt_bbdw.bbdw.dim_education b left join bbinfinity.dbo.usr_unc_highschoolinfo_ext x on b.educationsystemid = x.id
where a.constituentdimid = b.constituentdimid
for xml path('')
), 1, 1, '') as degree

from bbinfinity_rpt_bbdw.bbdw.dim_education a
join bbinfinity_rpt_bbdw.bbdw.dim_constituent c on a.constituentdimid = c.constituentdimid
join #pids_gaa_catalog p on c.constituentlookupid = p.val
),

-- board
board as
(
select distinct c.constituentdimid,

stuff
((
select distinct cast('$' as varchar(max)) + rtrim(ltrim(rl.role +
+ ' (' + case when r.startdate is null then '' else cast(convert(date,r.startdate) as varchar(50)) end +  
		case when r.enddate is null then '' else ' - ' + cast(convert(date,r.enddate) as varchar(50)) end + ')' +
		' ' + bx.rolecomments))
 
		from bbinfinity_rpt_bbdw.bbdw.fact_constituentgroupmember b
		left join bbinfinity_rpt_bbdw.bbdw.FACT_CONSTITUENTGROUPMEMBERROLE r on r.CONSTITUENTGROUPMEMBERFACTID = m.CONSTITUENTGROUPMEMBERFACTID
		left join bbinfinity_rpt_bbdw.bbdw.DIM_CONSTITUENTGROUPMEMBERROLE rl on rl.CONSTITUENTGROUPMEMBERROLEDIMID = r.CONSTITUENTGROUPMEMBERROLEDIMID
		left join BBInfinity.dbo.USR_UNC_UNCGROUPMEMBERROLE_EXTENSION bx on bx.id = r.constituentgroupmemberrolesystemid

where m.MEMBERCONSTITUENTDIMID = b.memberconstituentdimid
for xml path('')
), 1, 1, '') as board
				from bbinfinity_rpt_bbdw.bbdw.fact_constituentgroupmember m

		join bbinfinity_rpt_bbdw.bbdw.dim_constituentgroup g on g.CONSTITUENTGROUPDIMID = m.CONSTITUENTGROUPDIMID
		join bbinfinity_rpt_bbdw.bbdw.dim_constituent grp on grp.CONSTITUENTDIMID = g.CONSTITUENTDIMID
		join bbinfinity_rpt_bbdw.bbdw.dim_constituent c on c.constituentdimid = m.memberconstituentdimid
		join #pids_gaa_catalog p on c.constituentlookupid = p.val
		where grp.CONSTITUENTLOOKUPID = '8-11351244' 
       ),
		
-- gaa membership level
gaa as
(
select distinct
fm.constituentdimid,
mp.membershipprogram + ': ' + mp.membershiplevel gaa_member_level

from
bbinfinity_rpt_bbdw.bbdw.fact_member fm
join bbinfinity_rpt_bbdw.bbdw.dim_membership dm on fm.membershipdimid = dm.membershipdimid
join bbinfinity_rpt_bbdw.bbdw.dim_membershipprogram mp on dm.membershipprogramdimid = mp.membershipprogramdimid
join bbinfinity_rpt_bbdw.bbdw.dim_constituent c on fm.constituentdimid = c.constituentdimid
join #pids_gaa_catalog p on c.constituentlookupid = p.val

where
mp.membershipprogram like '%gaa%'
and dm.membershipstatusdimid = 1
and (dm.expirationdate >= getdate() or dm.expirationdate is null)
) 


select
p.sortorder,
d.constituentdimid,
d.constituentlookupid,
d.constituenttitle + ' ' + d.constituentfullname + ' ' + d.constituentsuffix fullname,
d.primaryaddress,
d.primaryaddresscity, 
d.primaryaddressstateabbreviation primaryaddressstate, 
d.primaryaddresspostcode, 
case when d.primaryaddresscountry = 'United States' then '' else d.primaryaddresscountry end as primaryaddresscountry,
d.primaryemailaddress,
d.primaryphonenumber,
gaa.gaa_member_level,
isnull(g.amount, 0) all_time_giving,
d.jobtitle,
d.employer,
e.degree,
b.board,
i.involvement educational_involvement,
case when dc.id is null then null else 'Yes' end as isdeceased,
substring(dc.deceaseddate,5,2) + '/' + substring(dc.deceaseddate,7,2) + '/' + substring(dc.deceaseddate,1,4) deceaseddate

from
bbinfinity_rpt_bbdw.bbdw.fact_usr_unc_demographic d
left join bbinfinity.dbo.deceasedconstituent dc on d.constituentsystemid = dc.id
left join bbinfinity_rpt_bbdw.bbdw.fact_usr_unc_aggregategiving_alltime g on d.constituentdimid = g.constituentdimid
left join involvement i on d.constituentdimid = i.constituentdimid
left join degree e on d.constituentdimid = e.constituentdimid
left join board b on b.constituentdimid = d.constituentdimid
left join gaa gaa on d.constituentdimid = gaa.constituentdimid
join #pids_gaa_catalog p on d.constituentlookupid = p.val

where 
d.isconstituent = 1

order by
p.sortorder

end

			