
declare @email nvarchar(100) = '**********@tmomail.net';
declare @cc nvarchar(100) = '*****@unc.edu';
declare @last_package nvarchar(200) = 'BBDW_FACT_USR_UNC_EDUCATIONALHISTORY.dtsx'
declare @etl_start_ind bit;
declare @etl_finished_ind bit;
declare @lastrun_package nvarchar(100);
declare @lastrun_package_start datetime;
declare @lastrun_package_runtime nvarchar(10);
declare @rundate date = convert(date, getdate());


set @etl_start_ind = 
(
select case when count(*) > 0 then 1 else 0 end 
from etlhistory where convert(date, etlstarttime) = @rundate and ssispackagename = 'bbdw_etl'
)

set @etl_finished_ind = 
(
select case when count(*) > 0 then 1 else 0 end 
from etlhistory where convert(date, etlstarttime) = @rundate and ssispackagename = @last_package
);

set @lastrun_package_start = 
(
select max(etlstarttime) as last_package_start
from etlhistory
where convert(date, etlstarttime) = @rundate
);

set @lastrun_package = 
(
select ssispackagename
from etlhistory
where convert(date, etlstarttime) = @rundate and etlstarttime = @lastrun_package_start
);

set @lastrun_package_runtime = (select convert(varchar, dateadd(second, (datediff(second, @lastrun_package_start, getdate()) ), 0), 108) as runtime);


select 
@email email, 
@cc cc,
case 
when @etl_start_ind = 0 then '** BBDW ETL DID NOT RUN **'
when @etl_finished_ind = 0 then '** BBDW ETL HAS NOT COMPLETED **'
else '' end as alert_subject,
'Last package: ' + @lastrun_package + '  |  Run Time: ' +  @lastrun_package_runtime as alert_text

where 
@etl_start_ind = 0 or @etl_finished_ind = 0

 