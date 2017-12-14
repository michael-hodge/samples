
options symbolgen;
options validvarname = any;
options validmemname = extend; 

* get list of delegate files in folder;
*%let dirpath = I:\Network Management Support\Development\Delegated Processing\UNC Delegated Processing\UNC Input files\;
%let dirpath=C:\Users\u101667\Desktop\unc_delegate\import\;

%macro get_delfiles(location);

	filename _dir_ "%bquote(&location.)";
	data DelFiles (keep = filepath);
	handle=dopen( '_dir_' );
		if handle > 0 then do;
	    	count=dnum(handle);
	    	do i=1 to count;
		  		filepath = "%bquote(&location.)" || dread(handle,i);
	      		output DelFiles;
	    		end;
	  	end;
	rc=dclose(handle);
	run;
	filename _dir_ clear;

%mend;

%get_delfiles(&dirpath)

* assign library reference to delgate files;
data DelLibraries ;
	set DelFiles ;
	length libref $8 rc 8 sysmsg $200 ;
	libref = 'DelLib' || put(_n_,Z2.);
	rc = libname(libref, filepath,,'access=readonly');
	if rc ne 0 then sysmsg = sysmsg();
run;

* create list of delegate files and sheets;
proc sql;
	create table DelSheets as 
	select * from dictionary.members where libname like 'DELLIB%' and memname not like '%FilterDatabase%';
quit;

* add sequence number to delegate dataset;
data DelSheets;
	set DelSheets;
	memname = strip(tranwrd(memname, "'", ""));
	id = _n_; 
run;

* get number of sheets;
proc sql noprint;
	select count(*) into :vSheetCount from DelSheets;
quit;

* create placeholder dataset;
data all_tabs;
input placeholder;
stop;
datalines;
1
;

* loop thru files/tabs listed in DelSheets to pull in records from linked delegate spreadsheets;
* create all_tabs dataset on first pass.  append additional records on subsequent passes;
%macro import(sequence_cnt);
   %do i = 1 %to &sequence_cnt;

		proc sql noprint;
			select count(*) into :vRecordCount from all_tabs;
		quit;

		data _null_;
			set DelSheets;
			call symput('id', libname);
			call symput('lib', libname);
			call symput('file', path);
			call symput('sheet', memname);
			call symput('sheet_literal', "'"||memname||"'n");
			where id = &i;
		run;

		 %if  &vRecordCount = 0  %then %do;

		 		proc sql;

					create table all_tabs as 
					select  
					today() as rundate format=mmddyys10.,
					symget('file') as delegate_file,
					tranwrd(symget('sheet'), '$', '') as delegate_sheet,
					*
					from &lib. .&sheet_literal.;

				quit;

		%end;

		%else %do;

		 		proc sql;

					create table all_tabs as 
					select * from all_tabs 
					union all
					select  
					today() as rundate format=mmddyys10.,
					symget('file') as delegate_file,
					tranwrd(symget('sheet'), '$', '') as delegate_sheet,
					*
					from &lib. .&sheet_literal.;

				quit;

		%end;

   %end;
%mend;

%import(&vSheetCount)

* create summary;
proc sql;
	create table summary as select delegate_file, delegate_sheet, count(*) as cnt from all_tabs group by delegate_file, delegate_sheet order by delegate_file, delegate_sheet;
quit;

* validation - compare column names and types across tabs.;
proc sql;
	create table columns as 
	select s.path, c.libname, c.memname, c.name, c.type, c.format 
	from dictionary.columns c, (select distinct libname, path from delsheets) s
	where c.libname = s.libname and c.libname like 'DELLIB%' and c.memname not like '%FilterDatabase%';
quit;

proc sql;
	create table column_counts as
	select name, count(*) as count from columns group by name order by name;
quit;

proc sql noprint;
	select max(count) into :vMaxColCount from (select name, count(*) as count from columns group by name);
quit;

proc sql;
	create table column_missing as select name, count from column_counts where count < &vMaxColCount;
quit;

proc sql;
create table validation as 
	select  'column appears in' || put(count, 2.) ||' out of ' || put(&vMaxColCount, 2.) || ' tabs' as error, name from column_counts where count < &vMaxColCount
union
	select 'column type mismatch' as error, a.name, a.memname as tabname1, b.memname as tabname2, a.type||':  '||a.format as format1, b.type||':  '||b.format as format2
	from columns a join columns b on a.libname = b.libname and a.name = b.name
	where (a.type <> b.type) or (a.format <> b.format);
quit;



