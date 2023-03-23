
with giving as
(
select constituentid, siteid, fy
from revenuesummary
where householdamount > 0
),

lapse as
(
select constituentid, siteid, fy, lag(fy, 1) over (partition by constituentid, siteid order by constituentid, siteid, fy) as lag
from giving
)

select 
c.id as CONSTITUENTID, 
s.id as SITEID, 
l.fy as FISCALYEAR, 
lag as PRIORYEAR, 
fy-lag as LAG,
case
when fy-lag = 1 then 'Retained'
when fy-lag between 2 and 5 then 'Reactivated (Short-lapsed)'
when fy-lag > 5 then 'Reactivated (Long-lapsed)'
else 'New Donor' 
end as STATUS

from
lapse l
join constituent c on l.constituentid = c.id
join sites s on s.id = l.siteid



