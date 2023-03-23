
select
mktsegment.id as MKTSEGMENTID,
mktsegment.name as MKTSEGMENTNAME,
mktsegmentation.id as MKTSEGMENTATIONID,
mktsegmentation.name as MKTSEGMENTATIONNAME,
mktsegmentation.siteid as SITEID,
mktsegment.segmenttype as SEGMENTTYPE,
mktsegmentation.description as DESCRIPTION,
mktsegmentation.activatedate as ACTIVATEDATE,
mktsegmentation.channel as CHANNEL,
mktsegmentation.maildate as LAUNCHDATE,
mktpackage.name as PACKAGENAME,
mktpackage.code as PACKAGECODE,
mktpackage.unitcost as PACKAGEUNITCOST,
mktpackage.costdistributionmethod as PACKAGEDISTRIBUTIONMETHOD,
mktpackage.channel as PACKAGECHANNEL,
mktpackagecategorycode.description as PACKAGECATEGORY,
parenttemplate.name as TEMPLATE,
sum(mktsegmentationsegmentactive.quantity) as QUANTITY,
sum(mktsegmentationsegmentactive.responders) as RESPONDERS,
sum(mktsegmentationsegmentactive.responses) as RESPONSES,
sum(mktsegmentationsegmentactive.variablecost) as VARIABLECOST,
sum(mktsegmentationsegmentactive.fixedcost) as FIXEDCOST,
sum(mktsegmentationsegmentactive.totalcost) as TOTALCOST,
sum(mktsegmentationsegmentactive.totalgiftamount) as TOTALGIFTAMOUNT,

cast((case 
when sum(mktsegmentationsegmentactive.totalgiftamount) > 0 
then sum(mktsegmentationsegmentactive.totalcost) / sum(mktsegmentationsegmentactive.totalgiftamount) 
else 0 end) as money) as COSTPERDOLLARRAISED,

cast((case 
when sum(mktsegmentationsegmentactive.responses) > 0 
then sum(mktsegmentationsegmentactive.totalgiftamount) / sum(mktsegmentationsegmentactive.responses) 
else 0 end) as money) as AVERAGEGIFTAMOUNT,

cast((case 
when sum(mktsegmentationsegmentactive.quantity) > 0 
then sum(cast(mktsegmentationsegmentactive.responses as decimal(19,4))) / sum(cast(mktsegmentationsegmentactive.quantity as decimal(19,4))) 
else 0 end) as decimal(19,4)) as RESPONSERATE,

cast((case 
when mktsegmentationactive.responserate > 0 
then cast((case 
when sum(mktsegmentationsegmentactive.quantity) > 0 
then (sum(cast(mktsegmentationsegmentactive.responses as decimal(19,4))) / sum(cast(mktsegmentationsegmentactive.quantity as decimal(19,4)))) * 100 
else 0 end) as decimal(19,4)) / mktsegmentationactive.responserate 
else 0 end) as decimal(19,4)) as LIFT,

sum(mktsegmentationsegmentactive.roiamount) as ROIAMOUNT,

cast((case 
when sum(mktsegmentationsegmentactive.totalcost) > 0 
then sum(mktsegmentationsegmentactive.roiamount) / sum(mktsegmentationsegmentactive.totalcost) 
else 0 end) as money) as ROIPERCENT,

cast((case 
when sum(mktsegmentationsegmentactive.expectedtotalgiftamount) > 0 
then sum(mktsegmentationsegmentactive.totalcost) / sum(mktsegmentationsegmentactive.expectedtotalgiftamount) 
else 0 end) as money) as EXPECTEDCOSTPERDOLLARRAISED,

sum(mktsegmentationsegmentactive.expectedresponders) as EXPECTEDRESPONDERS,
(mktsegmentationsegment.giftamount + isnull((select sum(giftamount) from dbo.mktsegmentationtestsegment where segmentid = mktsegmentationsegment.id),0)) / (1 + (select count(1) from dbo.mktsegmentationtestsegment where segmentid = mktsegmentationsegment.id)) as EXPECTEDGIFTAMOUNT,
sum(mktsegmentationsegmentactive.expectedtotalgiftamount) as EXPECTEDTOTALGIFTAMOUNT,
((mktsegmentationsegment.responserate + isnull((select sum(responserate) from dbo.mktsegmentationtestsegment where segmentid = mktsegmentationsegment.id),0)) / (1 + (select count(1) from dbo.mktsegmentationtestsegment where segmentid = mktsegmentationsegment.id))) / 100 as EXPECTEDRESPONSERATE,
sum(mktsegmentationsegmentactive.expectedroiamount) as EXPECTEDROIAMOUNT,

cast((case
when sum(mktsegmentationsegmentactive.totalcost) > 0 
then sum(cast(mktsegmentationsegmentactive.expectedroiamount as decimal(19,4))) / sum(cast(mktsegmentationsegmentactive.totalcost as decimal(19,4))) 
else 0 end) as decimal(19,4)) as EXPECTEDROIPERCENT

from 
dbo.mktsegmentationsegment
join dbo.mktsegmentation on mktsegmentation.id = mktsegmentationsegment.segmentationid
join dbo.mktsegmentationsegmentactive on mktsegmentationsegmentactive.segmentid = mktsegmentationsegment.id
join dbo.mktsegmentationactive on mktsegmentationactive.id = mktsegmentation.id
left join dbo.mktpackage on mktpackage.id = mktsegmentationsegment.packageid
left join dbo.mktsegment on mktsegment.id = mktsegmentationsegment.segmentid
left join dbo.mktpackagecategorycode on mktpackage.packagecategorycodeid = mktpackagecategorycode.id
left join dbo.mktcommunicationtemplate on mktcommunicationtemplate.mktsegmentationid = mktsegmentation.id
left join dbo.mktcommunicationtemplate as parenttemplate on parenttemplate.id = mktcommunicationtemplate.parentcommunicationtemplateid

where 
mktsegmentation.active = 1
and mktsegmentation.siteid != '93865C15-695A-4E49-B8D0-6A013936A338'

group by 
mktsegment.id, 
mktsegment.name, 
mktsegmentation.siteid,
mktsegment.segmenttype,
mktsegmentation.id,
mktsegmentation.name, mktsegmentation.description, 
mktsegmentation.activatedate, 
mktsegmentation.activatedate,
mktsegmentation.channel, 
mktsegmentationsegment.id, 
mktsegmentationsegment.packageid, 
mktsegmentationsegment.giftamount, 
mktsegmentationsegment.responserate,
mktpackage.channel, 
mktsegmentation.maildate,
mktpackagecategorycode.description, 
mktpackage.costdistributionmethod, 
mktpackage.code, 
mktpackage.name, 
mktpackage.unitcost, 
mktsegmentationactive.responserate,
parenttemplate.name 



