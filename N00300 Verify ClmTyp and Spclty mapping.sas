

libname srclib  "I:\Network Management Support\Development\Quality Audit Reporting\NPM Extracts\SASExtractfiles";
libname baselib "I:\Network Management Support\Development\Quality Audit Reporting\NPM Extracts\SASBasefiles";
libname mhslib "I:\Network Management Support\Development\Quality Audit Reporting\SASfilesGrid";

options ls=90 nosymbolgen nomprint mlogic mtrace missing=' '  mstored sasmstore=mymlib noxwait;

* get exclusion records;
proc import out = exclude 
datatable = 'Q00300_Input' 
dbms = access replace;
database = 'I:\Network Management Support\Development\Quality Audit Reporting\Audit Query Inputs\PowerMHS Audit Query Inputs.accdb';
usedate = yes;
scantime = no;
dbsaslabel = none;
run;

proc sql;

create table all_issues as

select
s.pracppn,
s.provppn,
v.provname,
v.prvtyp,
sp.spcltytier2,
v.specd1 as prac_specd1,
v.specd2 as prac_specd2,
v_prov.specd1 as prov_specd1,
s.network,
s.claim_type,
s.hold_code,
s.pay_code,
v.prov_userid as pr1011_userid,
v.prov_username as pr1011_username,
v.prov_lstchg_date as pr1011_lstchg_date,
s.userid as pr1032_userid,
s.username as pr1032_username,
s.lstchg_date as pr1032_lstchg_date,
case 
when v.specd1 in ('CHIRO') and s.claim_type not in ('CH') then 'specd1 CHIRO, claim_type not CH'
when v.specd1 in ('CHIRO') and s.hold_code in ('PL') and s.pay_code in ('S') then 'specd1 CHIRO, hold_code PL, pay_code S'
when v.specd1 in ('DME') and s.claim_type not in ('ME') then 'specd1 DME, claim_type not ME'
when v.specd1 in ('HH') and s.claim_type not in ('HH') then 'specd1 HH, claim_type not HH'
when sp.spcltytier2 in ('3600','3700') and v.specd1 not in ('NPMH') and s.claim_type not in ('MH') and v_prov.specd1 not in ('UCARE') then 'specd1 MH, claim_type not MH'
when v.specd1 not in ('CHIRO') and s.claim_type in ('CH') then 'specd1 not CHIRO, claim_type CH'
when v.specd1 in ('VOID') and s.end_date < '31DEC9999'd then 'specd1 VOID, active on pr1032'
when v.specd1 in ('VOID') and s.pracppn||s.provppn||s.network in (select pracppn||provppn||network from srclib.all_cnhmas_curr_info where end_date <> '31DEC9999'd) then 'specd1 VOID, active on pr1013'
when v.specd1 in ('UCARE') and s.claim_type not in ('UR') then 'specd1 UCARE, claim type not UR'
else '' end as reason_flag,
case when v.prov_lstchg_date >= s.lstchg_date then v.prov_userid  else s.userid end as userid,
case when v.prov_lstchg_date >= s.lstchg_date then v.prov_username else s.username end as username,
case when v.prov_lstchg_date >= s.lstchg_date then v.prov_lstchg_date else s.lstchg_date end as lstchg_date format=mmddyys10.,
compress(s.pracppn||s.provppn||v.provname||v.prvtyp||sp.spcltytier2||v.specd1||v.specd2||s.claim_type||s.hold_code||s.pay_code) as record_key

from 
srclib.all_psamas_curr_info s join srclib.all_provPR1011_info v on s.pracppn = v.providerppn join mhslib.spclty sp on v.specd1 = sp.spclty
join srclib.all_provPR1011_info v_prov on s.provppn = v_prov.providerppn join mhslib.spclty sp_prov on v_prov.specd1 = sp_prov.spclty

where
(v.specd1 in ('CHIRO') and s.claim_type not in ('CH'))
or (v.specd1 in ('CHIRO') and s.hold_code in ('PL') and s.pay_code in ('S'))
or (v.specd1 in ('DME') and s.claim_type not in ('ME'))
or (v.specd1 in ('HH') and s.claim_type not in ('HH'))
or (sp.spcltytier2 in ('3600','3700') and v.specd1 not in ('NPMH') and s.claim_type not in ('MH') and  v_prov.specd1 not in ('UCARE'))
or (v.specd1 not in ('CHIRO') and s.claim_type  in ('CH'))
or (v.specd1 in ('VOID') and s.end_date < '31DEC9999'd)
or (v.specd1 in ('VOID') and s.pracppn||s.provppn||s.network in (select pracppn||provppn||network from srclib.all_cnhmas_curr_info where end_date <> '31DEC9999'd)
or (v.specd1 in ('UCARE') and s.claim_type not in ('UR')))

order by
record_key;

quit;

* group concat networks;
data all_issues_network_concat (drop=network record_key);
     set all_issues;
     by record_key;
     retain networks;
     length networks $ 500;
     if first.record_key then networks = '';
     networks = catx('; ',trim(networks),network);
if last.record_key then output;
run;

* split out based on reason flag and exclude access records;
proc sql;

create table chiro as 
select today() as rundate format=mmddyys10., 'N00300' as qryName, 'Chiro' as Tabname, a.* 
from  all_issues_network_concat a
left join exclude e on a.provppn = e.provppn and a.pracppn = e.pracppn and lower(e.tabname) = 'chiro'
where reason_flag in ('specd1 CHIRO, claim_type not CH',  'specd1 CHIRO, hold_code PL, pay_code S', 'specd1 not CHIRO, claim_type CH') and a.provppn ne e.provppn and a.pracppn ne e.pracppn;

create table dme as 
select today() as rundate format=mmddyys10., 'N00300' as qryName, 'DME' as Tabname, a.* 
from  all_issues_network_concat a
left join exclude e on a.provppn = e.provppn and a.pracppn = e.pracppn and lower(e.tabname) = 'dme'
where reason_flag in ('specd1 DME, claim_type not ME') and a.provppn ne e.provppn and a.pracppn ne e.pracppn;

create table homehealth as 
select today() as rundate format=mmddyys10., 'N00300' as qryName, 'HomeHealth' as Tabname, a.* 
from  all_issues_network_concat a
left join exclude e on a.provppn = e.provppn and a.pracppn = e.pracppn and lower(e.tabname) = 'home health'
where reason_flag in ('specd1 HH, claim_type not HH') and a.provppn ne e.provppn and a.pracppn ne e.pracppn;

create table mentalhealth as 
select today() as rundate format=mmddyys10., 'N00300' as qryName, 'MentalHealth' as Tabname, a.* 
from  all_issues_network_concat a
left join exclude e on a.provppn = e.provppn and a.pracppn = e.pracppn and lower(e.tabname) = 'mental health'
where reason_flag in ('specd1 MH, claim_type not MH') and a.provppn ne e.provppn and a.pracppn ne e.pracppn;

create table pr1013voids as 
select today() as rundate format=mmddyys10., 'N00300' as qryName, 'PR1013Voids' as Tabname, a.* 
from  all_issues_network_concat a
left join exclude e on a.provppn = e.provppn and a.pracppn = e.pracppn and lower(e.tabname) = 'pr1013 voids'
where reason_flag in ('specd1 VOID, active on pr1013') and a.provppn ne e.provppn and a.pracppn ne e.pracppn;

create table pr1032voids as 
select today() as rundate format=mmddyys10., 'N00300' as qryName, 'PR1032Voids' as Tabname, a.* 
from  all_issues_network_concat a
left join exclude e on a.provppn = e.provppn and a.pracppn = e.pracppn and lower(e.tabname) = 'pr1032 voids'
where reason_flag in ( 'specd1 VOID, active on pr1032') and a.provppn ne e.provppn and a.pracppn ne e.pracppn;

create table urgentcare as 
select today() as rundate format=mmddyys10., 'N00300' as qryName, 'UrgentCare' as Tabname, a.* 
from  all_issues_network_concat a
left join exclude e on a.provppn = e.provppn and a.pracppn = e.pracppn and lower(e.tabname) = 'urgent care'
where reason_flag in ('specd1 UCARE, claim type not UR') and a.provppn ne e.provppn and a.pracppn ne e.pracppn;

quit;

 *** FOR TESTING: run interactively the next three statements interactively - BE SURE TO COMMENT THEM OUT FOR PRODUCTION ***;
* %include 'I:\Network Management Support\Development\WIP\Export for NPM queries_Test Version.sas';
* %let filepath=H:\Testrun\;
* %let filename=N00300.sas;

%Exportz(chiro, dme, homehealth, mentalhealth, urgentcare, pr1013voids, pr1032voids)

proc datasets library=work nolist kill;
run;
quit;
