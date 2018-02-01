/* Download Monthly NPI zip file from CMS website.  extract npi csv file and save as sas dataset */


/*libname nppeslib 'C:\nppes\';*/

libname nppeslib 'I:\Network Management Support\Development\Delegated Processing\NPPES\';

/* define variables*/
%let sysmonth = %sysfunc(month("&sysdate"d),monname9.);
%let sysyear= %sysfunc(year("&sysdate"d));
%let zippath = http://download.cms.gov/nppes/NPPES_Data_Dissemination_&sysmonth._&sysyear..zip;

/* download zip file from cms website*/
filename npi url "&zippath";
filename copy 'C:\nppes\nppes.zip';

data _null_;
	n=-1;
	infile npi recfm=s nbyte=n length=len;
	file copy recfm=n;
	input;
	put _infile_ $varying32767. len;
run;

/* get contents of zip file*/
filename inzip zip "C:\nppes\nppes.zip";

data contents(keep=memname);
	length memname $200;
	fid=dopen("inzip");
	if fid=0 then
		stop;
	memcount=dnum(fid);
	do i=1 to memcount;
	memname=dread(fid,i);
	output;
	end;
	rc=dclose(fid);
run;

/* assign npi filename to variable*/
proc sql noprint;
	select distinct btrim(memname) into :vNPIFileName from contents where memname like 'npidata_%.csv' and memname not like '%Header%';
quit;

/* import and save npi file*/
filename zipfile zip 'C:\nppes\nppes.zip';

data nppes;

	informat npi $10.;
	informat entity_type $1.;
	informat replacement_npi $10.;
	informat employer_identification_number $9.;
	informat organization_name $70.;
	informat last_name $35.;
	informat first_name $20.;
	informat middle_name $20.;
	informat name_prefix_text $5.;
	informat name_suffix_text $5.;
	informat credential_text $20.;
	informat other_organization_name $70.;
	informat other_organization_name_type $1.;
	informat other_last_name $35.;
	informat other_first_name $20.;
	informat other_middle_name $20.;
	informat other_name_prefix_text $5.;
	informat other_name_suffix_text $5.;
	informat other_credential_text $20.;
	informat other_last_name_type $1.;
	informat mailing_addr1 $55.;
	informat mailing_addr2 $55.;
	informat mailing_addr_city_name $40.;
	informat mailing_addr_state_name $40.;
	informat mailing_addr_zip $20.;
	informat mailing_addr_country_code $2.;
	informat mailing_addr_phone $20.;
	informat mailing_addr_fax_number $20.;
	informat practice_loc_addr1 $55.;
	informat practice_loc_addr2 $55.;
	informat practice_loc_addr_city_name $40.;
	informat practice_loc_addr_state_name $40.;
	informat practice_loc_addr_zip $20.;
	informat practice_loc_addr_country_code $2.;
	informat practice_loc_addr_phone $20.;
	informat practice_loc_addr_fax_number $20.;
	informat enumeration_date mmddyy10.;
	informat last_update_date mmddyy10.;
	informat deactivation_reason_code $2.;
	informat deactivation_date mmddyy10.;
	informat reactivation_date mmddyy10.;

	format npi $10.;
	format entity_type $1.;
	format replacement_npi $10.;
	format employer_identification_number $9.;
	format organization_name $70.;
	format last_name $35.;
	format first_name $20.;
	format middle_name $20.;
	format name_prefix_text $5.;
	format name_suffix_text $5.;
	format credential_text $20.;
	format other_organization_name $70.;
	format other_organization_name_type $1.;
	format other_last_name $35.;
	format other_first_name $20.;
	format other_middle_name $20.;
	format other_name_prefix_text $5.;
	format other_name_suffix_text $5.;
	format other_credential_text $20.;
	format other_last_name_type $1.;
	format mailing_addr1 $55.;
	format mailing_addr2 $55.;
	format mailing_addr_city_name $20.;
	format mailing_addr_state_name $20.;
	format mailing_addr_zip $20.;
	format mailing_addr_country_code $2.;
	format mailing_addr_phone $20.;
	format mailing_addr_fax_number $20.;
	format practice_loc_addr1 $55.;
	format practice_loc_addr2 $55.;
	format practice_loc_addr_city_name $40.;
	format practice_loc_addr_state_name $40.;
	format practice_loc_addr_zip $20.;
	format practice_loc_addr_country_code $2.;
	format practice_loc_addr_phone $20.;
	format practice_loc_addr_fax_number $20.;
	format enumeration_date mmddyy10.;
	format last_update_date mmddyy10.;
	format deactivation_reason_code $2.;
	format deactivation_date mmddyy10.;
	format reactivation_date mmddyy10.;

  	infile zipfile (%trim(&vNPIFileName))  delimiter = ',' missover dsd lrecl = 32767 firstobs = 2;
  	input npi $ entity_type $ replacement_npi $ employer_identification_number $ organization_name $ last_name $ first_name $ middle_name $ name_prefix_text $ name_suffix_text $ credential_text $ other_organization_name $ other_organization_name_type $ other_last_name $ other_first_name $ other_middle_name $ other_name_prefix_text $ other_name_suffix_text $ other_credential_text $ other_last_name_type $ mailing_addr1 $ mailing_addr2 $ mailing_addr_city_name $ mailing_addr_state_name $ mailing_addr_zip $ mailing_addr_country_code $ mailing_addr_phone $ mailing_addr_fax_number $ practice_loc_addr1 $ practice_loc_addr2 $ practice_loc_addr_city_name $ practice_loc_addr_state_name $ practice_loc_addr_zip $ practice_loc_addr_country_code $ practice_loc_addr_phone $ practice_loc_addr_fax_number $ enumeration_date last_update_date deactivation_reason_code $ deactivation_date reactivation_date;

run;

data nppeslib.nppes;
	set nppes;
run;

/* copy zip file to shared folder*/
options noxwait;
%sysexec copy "C:\nppes\nppes.zip" "I:\Network Management Support\Development\Delegated Processing\NPPES\nppes.zip" ;
