/*
title: weekly provider sampling report
purpose:  identify 30 random provider records that have been touched in the last 7 days and send to justin's team for audit purposes.
*/

options symbolgen ls = 120 mlogic mtrace mprint missing = ' ' mstored sasmstore = mymlib noxwait ;
libname srclib  'I:\Network Management Support\Development\Quality Audit Reporting\NPM Extracts\SASExtractfiles';

%include 'H:\autoexec.sas';

* connect to grid;
options connectpersist;
options connectwait;
options ls = 126 ps = 41 mprint mlogic symbolgen;

signon mynode;
rsubmit;

%let user1 = '******';
%let pw1 = '******';

* get ppns touched in last 7 days from blue provider;
proc sql;
   connect to db2 as db1 (user=&user1 password=&pw1 database=CLPRAP);
   create table touched_ppns as
		select *
		from connection to db1
			(select  prov_nbr, max(mod_dt) as mod_dt
			from nmo.audit_case_mstr
			where current_date - mod_dt <= 7
			group by prov_nbr);
   disconnect from db1;
quit;

proc download 
	extendsn = yes v6transport 
	data = touched_ppns  
	out = work.touched_ppns;
run;

endrsubmit;
signoff;

* get npis.  conatenate multiples into one field.;

/* pracs */
proc sql;

create table prac_npi as

select n.providerppn, btrim(n.npi_number) ||  '(' || btrim(n.npi_type) || ')' as npi, i.accept_patients
from srclib.prac_npi n, srclib.pracpr1011_info i
where n.providerppn = i.providerppn and n.expiration_date = '12-31-2999'
order by n.providerppn, n.npi_type, n.npi_number;

quit;

data prac_npi_concat (drop=npi);
	set prac_npi;
	by providerppn;
	retain npi_concat;
	length npi_concat $ 500;
	if first.providerppn then npi_concat = '';
	npi_concat = catx('; ',trim(npi_concat),npi);
if last.providerppn then output;
run;

/* provs */
proc sql;

create table prov_npi as

select n.providerppn, btrim(n.npi_number) ||  '(' || btrim(n.npi_type) || ')' as npi,  i.accept_patients
from srclib.prov_npi n, srclib.provpr1011_info i
where   n.providerppn = i.providerppn and n.expiration_date = '12-31-2999'
order by n.providerppn, n.npi_type, n.npi_number;

quit;

data prov_npi_concat (drop=npi);
	set prov_npi;
	by providerppn;
	retain npi_concat;
	length npi_concat $ 500;
	if first.providerppn then npi_concat = '';
	npi_concat = catx('; ',trim(npi_concat),npi);
	if last.providerppn then output;
run;

* get prac and prov records.. apply directory filtering logic;
proc sql;

create table all_records as

/* pracs */
select
nprac.npi_concat as pracnpi, nprov.npi_concat as provnpi, l.practitionerppn as pracppn, l.providerppn as provppn, trim(l.first_name)||' '||substr(trim(l.middle_name),1,1)||case when l.middle_name is not null then '. ' else '' end||trim(l.last_name) as pracname, l.provname, nprac.accept_patients, l.addr1, l.addr2, l.city, l.state, l.zip, l.phone

from
srclib.prac_all_locations l, srclib.pracppns p, prac_npi_concat  nprac, prov_npi_concat nprov

where
l.practitionerppn = p.providerppn
and l.practitionerppn = nprac.providerppn
and l.providerppn = nprov.providerppn
  and l.primary = 'Y'
  and l.print_in_directory = 'Y'
  and p.prvtyp not in ('EV','GF','GP','SP', 'GROUP')
  and l.speclty not in ('AMB', 'AMB*', 'ANES', 'BCNCI', 'CCMED', 'ERPHY', 'HBPP', 'HOSPP', 'INLAB', 'MONLY', 'NEOP*', 'NEOPE', 'PATH', 'PEDCC', 'PEDEM', 'RAD', 'SMEIN', 'SMEPR',  'VAMIL', 'VOID')
  and l.speclty not like 'AMB%' and l.speclty not like 'NEOP%'
  and p.prvtyp  not in ('IP','RX')
  and trim(l.practitionerppn)||trim(l.providerppn) in (select distinct trim(pracppn)||trim(provppn) from srclib.pracnetcycleattrib where network in ('PPRN') and participation_code in ('PP','PC','PF'))

union

/* provs */
select
nprov.npi_concat as pracnpi, nprov.npi_concat as provnpi, l.providerppn as pracppn, l.providerppn as provppn, l.provname as pracname, l.provname, nprov.accept_patients, l.addr1, l.addr2, l.city, l.state, l.zip, l.phone

from
srclib.prov_all_locations l, srclib.provppns p, prov_npi_concat nprov

where
l.providerppn = p.providerppn
and l.providerppn = nprov.providerppn
  and l.primary = 'Y'
  and l.print_in_directory = 'Y'
  and l.provname not like '%Unknown%'
  and p.prvtyp not in ('EV','GF','GP','SP')
  and l.speclty not in  ('AMB', 'AMB*', 'ANES', 'BCNCI', 'CCMED', 'ERPHY', 'HBPP', 'HOSPP', 'INLAB', 'MONLY', 'NEOP*', 'NEOPE',  'PATH', 'PEDCC', 'PEDEM', 'RAD', 'SMEIN', 'SMEPR', 'VAMIL', 'VOID')
  and l.speclty not like 'AMB%' and l.speclty not like 'NEOP%'
  and p.prvtyp  not in ('IP','RX')
  and trim(l.providerppn) in (select distinct trim(provppn) from srclib.provnetcycleattrib where network in ('PPRN') and participation_code in ('PP','PC','PF'));

quit;

* get random sample of 50 records;
proc sql outobs = 50;

create table sample_records(drop = rand) as

select 
ranuni(0) as rand,  a.pracppn, a.pracnpi, a.pracname, a.provppn, a.provnpi, a.provname as provname, a.accept_patients, a.addr1, a.addr2, a.city, a.state, a.zip, a.phone, t.mod_dt as mod_dt

from 
all_records a, work.touched_ppns t

where 
a.pracppn = t.prov_nbr

order by 
rand;

quit;

* connect to grid;
options connectpersist;
options connectwait;
options ls = 126 ps = 41 mprint mlogic symbolgen;

signon mynode;
rsubmit;

* put data set for XL on unix;
proc upload 
	data = sample_records 
	out = sample_records; 
run;

proc export
	data = sample_records
	outfile = "/sas_oth_grp/nm_dev_supp/reports/&sysdate. provider_sample_records.xlsx"
	dbms = xlsx
	replace;
run;

* send email;
%macro send_sample_records;
    options emailsys = smtp  emailhost = smtp.bcbsnc.com EMAILPORT = 25;

    data _null_;
    call symputx('exmth',put(intnx('month',today(), -1),monyy7.));
    run;
    %put &exmth.;
    filename mail email
    to = ('NCPrIMO.SystemSupport@bcbsnc.com','Mary.Hatley@bcbsnc.com', 'Erica.Ryder@bcbsnc.com', 'Tyler.Kaufman@bcbsnc.com', 'Justin.Bright@bcbsnc.com', 'Shatifa.Searles@bcbsnc.com', 'Bryce.Tobul@bcbsnc.com')
    from = ('NCPrIMO.SystemSupport@bcbsnc.com')
	cc = ('Michael.Hodge@bcbsnc.com')
    subject = 'weekly provider sampling report'
    attach = ("/sas_oth_grp/nm_dev_supp/reports/&sysdate. provider_sample_records.xlsx" content_type = "excel");
    data _null_;
    file mail;
      if _n_= 1 then do;
    Put 'attached:  sample of provider records touched in the last seven days.';
    end;
run;
%mend;

%send_sample_records;

endrsubmit;
signoff;
