
options symbolgen ls = 120 mlogic mtrace mprint missing = ' ' mstored sasmstore = mymlib noxwait ;
libname srclib  'I:\Network Management Support\Development\Quality Audit Reporting\NPM Extracts\SASExtractfiles';
libname local  'H:\';

%include 'H:\autoexec.sas';

* connect to grid;
options connectpersist;
options connectwait;
options ls = 126 ps = 41 mprint mlogic symbolgen;

signon mynode;
rsubmit;

%let user1 = '******';
%let pw1 = '******';

* get blue provider case and party data;
proc sql;
   connect to db2 as db1 (user=&user1 password=&pw1 database=CLPRAP);
   create table casedata as
		select *
		from connection to db1

		(select distinct
		cntl.audit_nm,
		cntl.audit_type_cd,
		cntl.from_dt,
		cntl.thru_dt,
		case.prov_nbr,
		field.mod_dt,
		case.anal_id

		from
		nmo.audit_bld_case_field field
		join nmo.audit_bld_case case on case.audit_bld_case_id = field.audit_bld_case_id
		join nmo.audit_bld_cntl cntl on cntl.audit_bld_cntl_id = case.audit_bld_cntl_id

		where
		year(cntl.from_dt) >= 2016);

	create table party as 
		select *
		from connection to db1

		(select * from nmo.party);

   disconnect from db1;
quit;

* get npm task data;
proc sql;
   connect to oracle as db1 (user=&user1 password=&pw1 path=nmpimp);
   create table taskdata as
		select *
		from connection to db1

		(select
		et.ds entity_type_desc,
		case when e.entity_type_id = 1 then provppn.ppn_number_uda else pracppn.ppn_number_uda end as ppn,
		case when e.entity_type_id = 1 then prov.name else prac.fname||' '||prac.lname end as name,
		td.datestamp taskdata_datestamp,
		trunc(td.datestamp) as taskdata_date,
		td.user_id taskdata_user_id,
		f.ds as field_name,
		td.value taskdata_value

		from
		portico.pv_tasks t
		join portico.pv_taskdata_text td on t.id = td.task_id
		join portico.pv_entities e on t.entity_id = e.id
		join portico.pv_entity_types et on e.entity_type_id = et.id
		join portico.pv_fields f on f.id = td.field_id
		left join portico.pp_prac prac on e.relation_id = prac.id
		left join portico.pp_prov prov on e.relation_id = prov.id
		left join portico.v_uda_prac_ppn pracppn on prac.id = pracppn.prac_id
		left join portico.v_uda_prov_ppn provppn on prov.id = provppn.prov_id

		where
		e.entity_type_id in (1,4)
		and td.field_id = 101040
		and td.datestamp >= to_date('2016-01-01', 'yyyy-mm-dd'));

   disconnect from db1;
quit;

proc download 
	extendsn = yes v6transport 
	data = casedata  
	out = work.casedata;
run;

proc download 
	extendsn = yes v6transport 
	data = party  
	out = work.party;
run;

proc download 
	extendsn = yes v6transport 
	data = taskdata  
	out = work.taskdata;
run;

endrsubmit;
signoff;

proc sql;

create table ProviderDBQA as

select distinct
c.audit_nm as audit_name,
p.full_name as username,
c.from_dt as start_date,
c.thru_dt as end_date,
c.prov_nbr as provider_number,
c.mod_dt as touched_date,
t.taskdata_value as sr_number,
today() as rundate format=mmddyys10.

from
casedata c
join party p on c.anal_id = p.id 
left join taskdata t on c.prov_nbr = t.ppn and c.mod_dt = datepart(t.taskdata_date)

order by
c.audit_nm, p.full_name, c.mod_dt, t.taskdata_value;

quit;

proc export
	data = work.ProviderDBQA
	outfile = "\\teamsite\dept\primo\Audit_Process and Procedures\ProviderDBQA.xlsx"
	dbms = excel replace;
	sheet = "data";
run;

proc datasets library=work nolist kill;
run;
quit;
