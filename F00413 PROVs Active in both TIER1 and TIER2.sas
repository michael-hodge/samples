
*create tier tables;
proc sql;

create table tier1 as
select n.*, r.prpr_name, r.prer_name, r.prpr_entity, r.prer_entity
from FCTflt.facets_networks_all n
join FCTflt.prpr_prer_relationships r on n.prpr_id = r.prpr_id
where n.nwnw_id in ('B00000000002') and n.nwpr_pfx in ('B002') and n.eff_date >= '01Jan2016'd;

create table tier2 as
select n.*, r.prpr_name, r.prer_name, r.prpr_entity, r.prer_entity
from FCTflt.facets_networks_all n
join FCTflt.prpr_prer_relationships r on n.prpr_id = r.prpr_id
where n.nwnw_id in ('B00000000002') and n.nwpr_pfx in ('B003')  and n.eff_date >= '01Jan2016'd;

quit;

*get records with overlapping effective dates;
proc sql;

create table  F00413H as

select
t1.provppn as ProvPPN,
t1.prer_name as ProvName,
t1.prer_entity as PRER_Entity,
t1.pracppn as PracPPN,
t1.prpr_name as PracName,
t1.prpr_id as PRPR_ID,
t1.prpr_entity as PRPR_Entity,
t1.par_in_network as TIER1_PAR,
t1.eff_date as TIER1_Eff_Date,
t1.term_date as TIER1_Term_Date,
t2.par_in_network as TIER2_PAR,
t2.eff_date as TIER2_Eff_Date,
t2.term_date as TIER2_Term_Date

from
tier1 t1, tier2 t2

where
t1.prpr_id = t2.prpr_id
and t1.prpr_term_date > today()
and (t1.par_in_network in ('Y') or t2.par_in_network in ('Y'))
and not(t1.eff_date > t2.term_date or t2.eff_date > t1.term_date)

order by
provppn, pracppn;

quit;

proc sql;
create table FCT_DUAL_TIER_Active as select today() as Rundate format = mmddyys10., 'F00413H' as QryName, 'FCT_DUAL_TIER_Active' as TabName,* from F00413H where TIER1_Term_Date >= today() and TIER2_Term_Date >= today();
create table FCT_DUAL_TIER_Inactive as select today() as Rundate format = mmddyys10., 'F00413H' as QryName, 'FCT_DUAL_TIER_Inactive' as TabName,* from F00413H where TIER1_Term_Date < today() or TIER2_Term_Date < today();
quit;

/*
*** EXPORT DATA TO EXCEL ***;
%include 'I:\Network Management Support\Development\Quality Audit Reporting\Facets Extracts\SAS Programs\Facets Queries\Test Version\Export for Facets queries.sas';
%let filepath=H:\Testrun\;
%let filename=F00413.sas;
*/
%ExportFz(FCT_DUAL_TIER_Active, FCT_DUAL_TIER_Inactive)

 proc datasets library=work nolist kill;
 run;
 quit;
