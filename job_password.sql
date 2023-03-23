


create view [dbo].[v_SSB_PASSWORD_EXP] as

select 
'BBCRM' as username, cast(expire_days as int) as expire_days, convert(varchar, dateadd(day, cast(expire_days as int), getdate()), 107) as expire_date 

from 
openquery([ssbintel.uncaa.unc.edu,1533],'select loginproperty (''****'', ''daysuntilexpiration'') as expire_days from sys.sql_logins where name = ''****'''); 

go



declare @expire_days int = (select top 1 expire_days from udodw.dbo.v_SSB_PASSWORD_EXP);
declare @expire_date varchar(50) = (select top 1 expire_date from udodw.dbo.v_SSB_PASSWORD_EXP);
declare @subject_text varchar(max) = 'SSB password expires in ' + cast(@expire_days as varchar(4)) + ' days'
declare @body_text varchar(max) = 'the password for BBCRM on ssbintel.uncaa.unc.edu will expire on:  ' + @expire_date


if @expire_days <= 5

begin

exec msdb.dbo.sp_send_dbmail  
    @profile_name = 'Mail',  
    @recipients = '*****@unc.edu',
    @copy_recipients = '*****@unc.edu',
    @subject = @subject_text,
    @body = @body_text,
    @importance = 'high'

end