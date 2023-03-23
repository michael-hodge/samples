
CREATE procedure [dbo].[etl_stats]
(
@level as varchar(10),
@date as date = null
)

as
begin

--declare @date as date;
--declare @level as varchar(10);
--set @date = getdate();
--set @level = 'detail';

set @date = isnull(@date, getdate());

select
x.package_name as package,
format(start_time, 'MM-dd-yyyy') as run_date,
case 
when execution_path like '%(1)%' then 1
when execution_path like '%(2)%' then 2
when execution_path like '%(3)%' then 3
when execution_path like '%(4)%' then 4
when execution_path like '%(5)%' then 5
else '' end as lane,
x.executable_name as step,
case
when xs.execution_result = 0 then 'ok'
when xs.execution_result = 1 then 'failed'
when xs.execution_result = 2 then 'completed'
when xs.execution_result = 3 then 'cancelled'
end as result,
xs.start_time,
xs.end_time,
convert(varchar(10), xs.execution_duration/(1000*60*60)) + ' hr ' 
+ convert(varchar(10), (xs.execution_duration%(1000*60*60))/(1000*60)) + ' min '
+ convert(varchar(10), ((xs.execution_duration%(1000*60*60))%(1000*60))/1000) + ' sec' as run_time

into 
#detail

from 
ssisdb.catalog.executables x
join ssisdb.catalog.executable_statistics xs on x.executable_id = xs.executable_id and x.execution_id = xs.execution_id

where 
x.package_name = 'package.dtsx'
and convert(date, start_time) = convert(date, @date);


if @level = 'detail'
begin
	select * 
	from #detail 
	where lane in (1,2,3,4,5)
	order by lane, start_time
end

else

begin
if @level = 'summary'
	select 
	package, run_date, convert(varchar(1), lane) as lane, count(*) as step_cnt, format(min(start_time),'hh:mm') as lane_start, format(max(end_time),'hh:mm') as lane_end, 
	convert(varchar(10), datediff(millisecond, min(start_time), max(end_time))/(1000*60*60)) + ' hr ' 
	+ convert(varchar(10), (datediff(millisecond, min(start_time), max(end_time))%(1000*60*60))/(1000*60)) + ' min '
	+ convert(varchar(10), ((datediff(millisecond, min(start_time), max(end_time))%(1000*60*60))%(1000*60))/1000) + ' sec' as run_time

	from 
	#detail

	where 
	lane in (1,2,3,4,5)

	group by 
	package, run_date, lane

	union

	select 
	min(package) package,
	min(run_date) rund_date,
	'Total' lane,
	count(*) step_cnt,
	format(min(start_time),'hh:mm') as lane_start, format(max(end_time),'hh:mm') as lane_end, 
	convert(varchar(10), datediff(millisecond, min(start_time), max(end_time))/(1000*60*60)) + ' hr ' 
	+ convert(varchar(10), (datediff(millisecond, min(start_time), max(end_time))%(1000*60*60))/(1000*60)) + ' min '
	+ convert(varchar(10), ((datediff(millisecond, min(start_time), max(end_time))%(1000*60*60))%(1000*60))/1000) + ' sec' as run_time

	from 
	#detail

else
	print 'select summary or detail'

end

end