delimiter //

create procedure iup_update()

begin


-- initialize point and reason fields
update
iup_sales

set 
ssd_pnt = 0,
sata_pnt = 0,
pcie_pnt = 0,
system_pnt = 0,
hero_pnt = 0,
wireless_docking_pnt = 0,
realsense_pnt = 0,
vpro_pnt = 0,
wireless_display_pnt = 0,
reason = '';


-- convert description to all caps
update
iup_sales

set
product_description = trim(upper(product_description));


-- create edited product description
update
iup_sales

set
product_description_edit = concat(oem_system_part_number, '    ', product_description),
product_description_edit = replace(product_description_edit, '-', '_');


-- -----------------------------------------------------------------------------
-- client systems
update
iup_sales
set
  client_ind = 1,
  system_pnt = 20,
  reason = concat(reason, '\n', 'System: Acer Aspire R13 (20)')
where
  match(product_description_edit) against ('+ACER +ASPIRE +R13' in boolean mode);
  
update
iup_sales
set
  client_ind = 1,
  system_pnt = 12,
  reason = concat(reason, '\n', 'System: Acer Chromebook C720 (12)')
where
  match(product_description_edit) against ('+ACER +CHROMEBOOK +C720*' in boolean mode);
  
update
iup_sales
set
  client_ind = 1,
  system_pnt = 12,
  reason = concat(reason, '\n', 'System: Acer Iconica Tab 8 (12)')
where
  match(product_description_edit) against ('(+ACER +ICONIA +8") +(TAB* NET_TAB*)' in boolean mode);
  
update
iup_sales
set
  client_ind = 1,
  system_pnt = 20,
  reason = concat(reason, '\n', 'System: Acer Travelmate P6 (20)')
where
  match(product_description_edit) against ('+ACER +TRAVELMATE +P6' in boolean mode);
  
update
iup_sales
set
  client_ind = 1,
  system_pnt = 20,
  reason = concat(reason, '\n', 'System: Acer Veriton 4 (20)')
where
  match(product_description_edit) against ('+ACER +VERITON' in boolean mode)
  and product_description_edit like '% 4 %';
  
update
iup_sales
set
  client_ind = 1,
  system_pnt = 20,
  reason = concat(reason, '\n', 'System: Acer VZ4810G i34350X (20)')
where
  match(product_description_edit) against ('+ACER +VZ4810G +i34350X' in boolean mode);
  
update
iup_sales
set
  client_ind = 1,
  system_pnt = 12,
  reason = concat(reason, '\n', 'System: Asus CHROMEBOOK C300 (12)')
where
  match(product_description_edit) against ('+ASUS +CHROMEBOOK +C300' in boolean mode);
  
update
iup_sales
set
  client_ind = 1,
  system_pnt = 20,
  reason = concat(reason, '\n', 'System: Asus EEE Slate B121 (20)')
where
  match(product_description_edit) against ('+ASUS +EEE +SLATE +B121' in boolean mode);
  
update
iup_sales
set
  client_ind = 1,
  system_pnt = 12,
  reason = concat(reason, '\n', 'System: Asus Memo Pad 8 (12)')
where
  match(product_description_edit) against ('+ASUS +MEMO +PAD' in boolean mode)
  and product_description_edit like '% 8 %';
  
update
iup_sales
set
  client_ind = 1,
  system_pnt = 12,
  reason = concat(reason, '\n', 'System: Asus Transformer T100 (12)')
where
  match(product_description_edit) against ('+ASUS +TRANSFORMER +T100*' in boolean mode);
  
update
iup_sales
set
  client_ind = 1,
  system_pnt = 12,
  reason = concat(reason, '\n', 'System: Asus VivoTab (12)')
where
  match(product_description_edit) against ('+ASUS +VIVOTAB' in boolean mode);
  
update
iup_sales
set
  client_ind = 1,
  system_pnt = 20,
  reason = concat(reason, '\n', 'System: Asus BP1AE (20)')
where
  match(product_description_edit) against ('+ASUS +BP1AE' in boolean mode);
  
update
iup_sales
set
  client_ind = 1,
  system_pnt = 20,
  reason = concat(reason, '\n', 'System: Asus ET2221IUTH (20)')
where
  match(product_description_edit) against ('+ASUS +ET2221IUTH' in boolean mode);
  
update
iup_sales
set
  client_ind = 1,
  system_pnt = 20,
  reason = concat(reason, '\n', 'System: Asus AsusPro BU401LA (20)')
where
  match(product_description_edit) against ('+ASUSPRO +BU401LA' in boolean mode);
  
update
iup_sales
set
  client_ind = 1,
  system_pnt = 20,
  reason = concat(reason, '\n', 'System: Asus AsusPro BU401LG (20)')
where
  match(product_description_edit) against ('+ASUSPRO +BU401LG' in boolean mode);
  
update
iup_sales
set
  client_ind = 1,
  system_pnt = 12,
  reason = concat(reason, '\n', 'System: Dell Chromebook 3010 (12)')
where
  match(product_description_edit) against ('+DELL +CHROMEBOOK +3010' in boolean mode);
  
update
iup_sales
set
  client_ind = 1,
  system_pnt = 20,
  reason = concat(reason, '\n', 'System: Dell Latitude 7000 (20)')
where
  match(product_description_edit) against ('+DELL +LATITUDE +E7*)' in boolean mode);
    
update
iup_sales
set
  client_ind = 1,
  system_pnt = 20,
  reason = concat(reason, '\n', 'System: Dell Optiplex 9020 Micro (20)')
where
  match(product_description_edit) against ('+DELL +OPTIPLEX +9020 +MICRO' in boolean mode);
  
update
iup_sales
set
  client_ind = 1,
  system_pnt = 20,
  reason = concat(reason, '\n', 'System: Dell Optiplex 9030 (20)')
where
  match(product_description_edit) against ('+DELL +OPTIPLEX +9030 ' in boolean mode);
  
update
iup_sales
set
  client_ind = 1,
  system_pnt = 20,
  reason = concat(reason, '\n', 'System: Dell Venue 11 Pro (20)')
where
  match(product_description_edit) against ('+DELL +VENUE +11 +PRO ' in boolean mode);
  
update
iup_sales
set
  client_ind = 1,
  system_pnt = 12,
  reason = concat(reason, '\n', 'System: Dell Venue 8 Pro (12)')
where
  match(product_description_edit) against ('+DELL +VENUE +PRO' in boolean mode)
  and product_description_edit like '% 8 %';
  
update
iup_sales
set
  client_ind = 1,
  system_pnt = 12,
  reason = concat(reason, '\n', 'System: Entegra Crossfire Pro (12)')
where
  match(product_description_edit) against ('+ENTEGRA +CROSSFIRE +PRO ' in boolean mode);
  
update
iup_sales
set
  client_ind = 1,
  system_pnt = 20,
  reason = concat(reason, '\n', 'System: Fujitsu Esprimo X923 (20)')
where
  match(product_description_edit) against ('+FUJITSU +ESPRIMO +X923' in boolean mode);
  
update
iup_sales
set
  client_ind = 1,
  system_pnt = 20,
  reason = concat(reason, '\n', 'System: Fujitsu Lifebook T725/T904/T935/U745 (20)')
where
  match(product_description_edit) against ('(+FUJITSU +LIFEBOOK) +(T725 T904 T935 U745)' in boolean mode);
  
update
iup_sales
set
  client_ind = 1,
  system_pnt = 12,
  reason = concat(reason, '\n', 'System: Fujitsu Stylistic Q355/Q555 (12)')
where
  match(product_description_edit) against ('(+FUJITSU +STYLISTIC) +(Q355 Q555)' in boolean mode);
  
update
iup_sales
set
  client_ind = 1,
  system_pnt = 20,
  reason = concat(reason, '\n', 'System: Google Chromebook Pixel (20)')
where
  match(product_description_edit) against ('+GOOGLE +CHROMEBOOK +PIXEL' in boolean mode);

update
iup_sales
set
  client_ind = 1,
  system_pnt = 12,
  reason = concat(reason, '\n', 'System: HP Chromebook 11 (12)')
where
  match(product_description_edit) against ('+HP +CHROMEBOOK +11' in boolean mode);
  
update
iup_sales
set
  client_ind = 1,
  system_pnt = 12,
  reason = concat(reason, '\n', 'System: HP Chromebook 14 (12)')
where
  match(product_description_edit) against ('+HP +CHROMEBOOK +14' in boolean mode);
  
update
iup_sales
set
  client_ind = 1,
  system_pnt = 20,
  reason = concat(reason, '\n', 'System: HP Elitebook Revolve 810 (20)')
where
  match(product_description_edit) against ('+HP +810* -8100' in boolean mode);
  
update
iup_sales
set
  client_ind = 1,
  system_pnt = 20,
  reason = concat(reason, '\n', 'System: HP Elitebook Folio 9480m (20)')
where
  match(product_description_edit) against ('(+9480m)' in boolean mode);  
  
update
iup_sales
set
  client_ind = 1,
  system_pnt = 20,
  reason = concat(reason, '\n', 'System: HP Elitebook Folio 1040 G1 (20)')
where
  match(product_description_edit) against ('+HP +1040 -10400 -9480*' in boolean mode);
  
update
iup_sales
set
  client_ind = 1,
  system_pnt = 20,
  reason = concat(reason, '\n', 'System: HP Elitebook Folio 1020 G1 (20)')
where
  match(product_description_edit) against ('+HP +1020 -10400 -9480*' in boolean mode);

update
iup_sales
set
  client_ind = 1,
  system_pnt = 20,
  reason = concat(reason, '\n', 'System: HP Elitedesk 800 G1 Mini (20)')
where
  match(product_description_edit) against ('(+HP +ELITEDESK +800 +G1 +(MINI MICRO SMALL TINY)' in boolean mode);
  
update
iup_sales
set
  client_ind = 1,
  system_pnt = 20,
  reason = concat(reason, '\n', 'System: HP Eliteone 800 (20)')
where
  match(product_description_edit) against ('(+HP +ELITEONE +800)' in boolean mode);
  
update
iup_sales
set
  client_ind = 1,
  system_pnt = 12,
  reason = concat(reason, '\n', 'System: HP Elitepad 1000 G2 (12)')
where
  match(product_description_edit) against ('(+HP +ELITEPAD +1000 +G2)' in boolean mode);
  
update
iup_sales
set
  client_ind = 1,
  system_pnt = 12,
  reason = concat(reason, '\n', 'System: HP Stream 11 Pro (12)')
where
  match(product_description_edit) against ('(+HP +Stream +11 +PRO)' in boolean mode);
  
update
iup_sales
set
  client_ind = 1,
  system_pnt = 20,
  reason = concat(reason, '\n', 'System: HP Elite x2 1011 (20)')
where
  match(product_description_edit) against ('(+HP +ELITE +X2 +1011)' in boolean mode);
  
update
iup_sales
set
  client_ind = 1,
  system_pnt = 12,
  reason = concat(reason, '\n', 'System: HP Pro Tablet EE 10 Education (12)')
where
  match(product_description_edit) against ('(+HP +PRO +TABLET +EE +10 +Education)' in boolean mode);
  
update
iup_sales
set
  client_ind = 1,
  system_pnt = 12,
  reason = concat(reason, '\n', 'System: HP Pro Tablet 408/610 G1 (12)')
where
  match(product_description_edit) against ('(+HP +PRO +TABLET +G1) +(408 610)' in boolean mode);
  
update
iup_sales
set
  client_ind = 1,
  system_pnt = 20,
  reason = concat(reason, '\n', 'System: HP Pro x2 612 (20)')
where
  match(product_description_edit) against ('(+HP +PRO +X2 +612)' in boolean mode);
  
update
iup_sales
set
  client_ind = 1,
  system_pnt = 20,
  reason = concat(reason, '\n', 'System: HP Z1 G2 (20)')
where
  match(product_description_edit) against ('(+HP +Z1 +G2)' in boolean mode);
  
update
iup_sales
set
  client_ind = 1,
  system_pnt = 20,
  reason = concat(reason, '\n', 'System: Lenovo ThinkCentre M93p/M93z/Tiny (20)')
where
  match(product_description_edit) against ('+(THINKCENTRE THINKCNTR LENOVO) +(M93P M93Z Tiny)' in boolean mode);
  
update
iup_sales
set
  client_ind = 1,
  system_pnt = 12,
  reason = concat(reason, '\n', 'System: Lenovo Thinkpad 11e (12)')
where
  match(product_description_edit) against ('(+THINKPAD +11E -YOGA)' in boolean mode);
  
update
iup_sales
set
  client_ind = 1,
  system_pnt = 20,
  reason = concat(reason, '\n', 'System: Lenovo Thinkpad Helix 2nd Gen (20)')
where
  match(product_description_edit) against ('+THINKPAD +HELIX +(G2 2ND)' in boolean mode);
  
update
iup_sales
set
  client_ind = 1,
  system_pnt = 20,
  reason = concat(reason, '\n', 'System: Lenovo Thinkpad T440s (20)')
where
  match(product_description_edit) against ('(+THINKPAD +T440S)' in boolean mode);
  
update
iup_sales
set
  client_ind = 1,
  system_pnt = 12,
  reason = concat(reason, '\n', 'System: Lenovo Thinkpad Tablet 10 (12)')
where
  match(product_description_edit) against ('+THINKPAD +TABLET +10' in boolean mode);
  
update
iup_sales
set
  client_ind = 1,
  system_pnt = 20,
  reason = concat(reason, '\n', 'System: Lenovo Thinkpad x1 Carbon (20)')
where
  match(product_description_edit) against ('+X1 -1G* -2G*' in boolean mode);
  
update
iup_sales
set
  client_ind = 1,
  system_pnt = 20,
  reason = concat(reason, '\n', 'System: Lenovo Thinkpad Yoga (20)')
where
  match(product_description_edit) against ('+YOGA* +(THINKPAD 11E)' in boolean mode);
  
update
iup_sales
set
  client_ind = 1,
  system_pnt = 12,
  reason = concat(reason, '\n', 'System: Lenovo N21 Chromebook (12)')
where
  match(product_description_edit) against ('+LENOVO +CHROMEBOOK +N21' in boolean mode);
  
update
iup_sales
set
  client_ind = 1,
  system_pnt = 20,
  reason = concat(reason, '\n', 'System: Microsoft Surface Pro 3 (20)')
where
  match(product_description_edit) against ('+(SURFACE MICROSOFTSURFACE SRFC)' in boolean mode)
  and product_description_edit like '% 3 %';

update
iup_sales
set
  client_ind = 1,
  system_pnt = 20,
  reason = concat(reason, '\n', 'System: Motion R12 (20)')
where
  match(product_description_edit) against ('+MOTION +R12' in boolean mode);
  
update
iup_sales
set
  client_ind = 1,
  system_pnt = 20,
  reason = concat(reason, '\n', 'System: Panasoic Toughpad FZ-G1 (20)')
where
  match(product_description_edit) against ('+PANASONIC +TOUGHPAD +"FZ_G1"' in boolean mode);
  
update
iup_sales
set
  client_ind = 1,
  system_pnt = 12,
  reason = concat(reason, '\n', 'System: Panasoic Toughpad FZ-B2/FZ-M1/FZ-R1 (12)')
where
  match(product_description_edit) against ('(+PANASONIC +TOUGHPAD) +("FZ_B2" "FZ_M1" "FZ_R1")' in boolean mode);
  
update
iup_sales
set
  client_ind = 1,
  system_pnt = 20,
  reason = concat(reason, '\n', 'System: Samsung ATIV 9 (20)')
where
  match(product_description_edit) against ('+ATIV' in boolean mode)
  and (product_description_edit like '% 9 %' or product_description_edit like '%BOOK9%');
  
update
iup_sales
set
  client_ind = 1,
  system_pnt = 12,
  reason = concat(reason, '\n', 'System: Toshiba Chromebook 2 (12)')
where
  match(product_description_edit) against ('+TOSHIBA +CHROMEBOOK' in boolean mode)
  and product_description_edit like '% 2 %';
  
update
iup_sales
set
  client_ind = 1,
  system_pnt = 12,
  reason = concat(reason, '\n', 'System: Toshiba Encore 2 (12)')
where
  match(product_description_edit) against ('+TOSHIBA +ENCORE' in boolean mode)
  and product_description_edit like '% 2 %';
  
update
iup_sales
set
  client_ind = 1,
  system_pnt = 20,
  reason = concat(reason, '\n', 'System: Toshiba Portege Z20T/Z30 (20)')
where
  match(product_description_edit) against ('(+TOSHIBA +PORTEGE) +(Z20T Z30)' in boolean mode);
  
update
iup_sales
set
  client_ind = 1,
  system_pnt = 20,
  reason = concat(reason, '\n', 'System: Toshiba Tecra Z40/Z50 (20)')
where
  match(product_description_edit) against ('(+TOSHIBA +TECRA) +(Z40 Z50)' in boolean mode);
  
update
iup_sales
set
  client_ind = 1,
  system_pnt = 12,
  reason = concat(reason, '\n', 'System: Xplore Bobcat (12)')
where
  match(product_description_edit) against ('+XPLORE +BOBCAT' in boolean mode);
  
update
iup_sales
set
  client_ind = 1,
  system_pnt = 12,
  reason = concat(reason, '\n', 'System: Xplore XC6 DMSR (12)')
where
  match(product_description_edit) against ('+XPLORE +XC6 +DMSR' in boolean mode);
  
update
iup_sales
set
  client_ind = 1,
  system_pnt = 12,
  reason = concat(reason, '\n', 'System: Standard Core Desktop (12)')
where
  reason = ''
  and match(product_description_edit) against  ('+(DESKTOP* WKSTN WORKSTATION* PRECISION PRODESK OPTIPLEX ELITEDESK THINKCENTRE "HP 600*" "HP 800*" "HP ED 800" 800ED "DELL 3020*" "DELL 3030*" "DELL 7020*" "DELL 9020*" "ALL-IN-ONE" NUC "ELITED DESK" "ELITE 800" AIO 800ED*) +(CORE I3* I5* I7*)' in boolean mode);


update
iup_sales
set
  client_ind = 1,
  system_pnt = 12,
  reason = concat(reason, '\n', 'System: Standard Unidentified Desktop (12)')
where
  reason = ''
  and match(product_description_edit) against ('+(DESKTOP* WKSTN WORKSTATION* PRECISION PRODESK OPTIPLEX ELITEDESK THINKCENTRE "HP 600*" "HP 800*" "HP ED 800" 800ED "DELL 3020*" "DELL 3030*" "DELL 7020*" "DELL 9020*" "ALL-IN-ONE" NUC "ELITED DESK" "ELITE 800" AIO 800ED*) -CELERON -PENTIUM' in boolean mode);


-- -----------------------------------------------------------------------------
-- client add-on
update
iup_sales
set
  vpro_pnt = 10,
  reason = concat(reason, '\n', 'Add-on: vPro (10)')
where
  client_ind = 1
  and match(product_description_edit) against ('VPRO' in boolean mode);

update
iup_sales
set
  realsense_pnt = 5,
  reason = concat(reason, '\n', 'Add-on: RealSense (10)')
where
  client_ind = 1
  and match(product_description_edit) against ('REALSENSE "REAL SENSE"' in boolean mode);

update
iup_sales
set
  wireless_docking_pnt = 5,
  reason = concat(reason, '\n', 'Add-on: Wireless Docking (10)')
where
  client_ind = 1
  and match(product_description_edit) against ('"WIRELESS DOCKING" "WI DI" WIDI' in boolean mode);
  
update
iup_sales
set
  ssd_pnt = 10,
  reason = concat(reason, '\n', 'Add-on: Client SSD (10)')
where
  client_ind = 1
  and match(product_description_edit) against ('SSD "SOLID STATE" "SOLID_STATE"' in boolean mode);
  
  
-- -----------------------------------------------------------------------------
-- client hero skus
update
iup_sales
set
  hero_pnt = 5,
  reason = concat(reason, '\n', 'Hero: Lenovo ThinkPad Yoga Ultrabook (5)')
where
  client_ind = 1
  and match(product_description_edit) against ('+YOGA* +(THINKPAD 11E)' in boolean mode);
  
update
iup_sales
set
  hero_pnt = 5,
  reason = concat(reason, '\n', 'Hero: Microsoft Surface Pro 3 (5)')
where
  client_ind = 1
  and match(product_description_edit) against ('+(SURFACE MICROSOFTSURFACE SRFC)' in boolean mode)
  and product_description_edit like '% 3 %';
  
update
iup_sales
set
  hero_pnt = 5,
  reason = concat(reason, '\n', 'Hero: Dell Latitude 13 7000 Series (5)')
where
  client_ind = 1
  and match(product_description_edit) against ('+DELL +LATITUDE +13 +E7*' in boolean mode);
  
update
iup_sales
set
  hero_pnt = 5,
  reason = concat(reason, '\n', 'Hero: Dell Latitude 14 7000 Series (5)')
where
  client_ind = 1
  and match(product_description_edit) against ('+DELL +LATITUDE +14 +E7*' in boolean mode);
  
update
iup_sales
set
  hero_pnt = 5,
  reason = concat(reason, '\n', 'Hero: HP Elite x2 1011 G1 (5)')
where
  client_ind = 1
  and match(product_description_edit) against ('+HP +ELITE +X2 +1011 +G1' in boolean mode);
  
-- removed 07/23
-- update
-- iup_sales
-- set
--   hero_pnt = 5,
--   reason = concat(reason, '\n', 'Hero: Lenovo Thinkpad Helix 2nd Gen (5)')
-- where
--   client_ind = 1
--   and match(product_description_edit) against ('+THINKPAD +HELIX +(G2 2ND)' in boolean mode);


-- removed 07/23  
-- update
-- iup_sales
-- set
--   hero_pnt = 5,
--   reason = concat(reason, '\n', 'Hero: Dell Venue 11 Pro 7000 (5)')
-- where
--   client_ind = 1
--  and match(product_description_edit) against ('+DELL +VENUE +E7* +11' in boolean mode);

update
iup_sales
set
  hero_pnt = 5,
  reason = concat(reason, '\n', 'Hero: Dell Optiplex 9020 Micro (5)')
where
  client_ind = 1
  and match(product_description_edit) against ('(+DELL +OPTIPLEX +9020 +MICRO)' in boolean mode);
  
update
iup_sales
set
  hero_pnt = 5,
  reason = concat(reason, '\n', 'Hero: HP Elitedesk 800 G1 Mini (5)')
where
  client_ind = 1
  and match(product_description_edit) against ('(+HP +ELITEDESK +800 +G1 +(MINI MICRO SMALL TINY)' in boolean mode);
  
update
iup_sales
set
  hero_pnt = 5,
  reason = concat(reason, '\n', 'Hero: Lenovo Thinkcentre M93p Tiny (5)')
where
  client_ind = 1
  and match(product_description_edit) against ('+THINKCENTRE +M93P +TINY' in boolean mode);
  
update
iup_sales
set
  hero_pnt = 5,
  reason = concat(reason, '\n', 'Hero: HP Elitebook Folio 1040 G2 (5)')
where
  client_ind = 1
  and match(product_description_edit) against ('(+HP +G2 +(FOLIO 1040)' in boolean mode);
  
update
iup_sales
set
  hero_pnt = 5,
  reason = concat(reason, '\n', 'Hero: HP Elitebook Folio 1040 G1 (5)')
where
  client_ind = 1
  and match(product_description_edit) against ('(+HP +G1 +(FOLIO 1040)' in boolean mode);
  
  
  
-- -----------------------------------------------------------------------------
-- server systems
update
iup_sales
set
  server_ind = 1,
  system_pnt = 5,
  reason = concat(reason, '\n', 'System: E3 Server (5)')
where
   reason = ''
   and match(product_description_edit) against ('E3_* -ASPIRE' in boolean mode);
   -- match(product_description_edit) against ('+RD* +E5*' in boolean mode);

update
iup_sales
set
  server_ind = 1,
  system_pnt = 15,
  reason = concat(reason, '\n', 'System: E5 Server (15)')
where
   reason = ''
   and match(product_description_edit) against ('E5_* -ASPIRE' in boolean mode);

update
iup_sales
set
  server_ind = 1,
  system_pnt = 45,
  reason = concat(reason, '\n', 'System: E7 Server (45)')
where
   reason = ''
   and match(product_description_edit) against ('E7_* -ASPIRE' in boolean mode);
   
   
update
iup_sales
set
  server_ind = 1,
  system_pnt = 5,
  reason = concat(reason, '\n', 'System: Unidentified Server (5)')
where
   reason = ''
   and match(product_description_edit) against ('SERV* SVR BLADE UCS PROLIANT POWEREDGE' in boolean mode);
   
-- -----------------------------------------------------------------------------
-- server add-ons
update iup_sales
set 
  sata_pnt = 10,
  reason = concat(reason, '\n', 'Add-on: SATA SSD (10)')
where
  server_ind = 1
  and match(product_description_edit) against ('SATA' in boolean mode);

update iup_sales
set 
  -- server_ind = 1,
  pcie_pnt = 25,
  reason = concat(reason, '\n', 'Add-on: PCIe SSD (25)')
where
  server_ind = 1
  and match(product_description_edit) against ('PCIE' in boolean mode);

update iup_sales
set 
  ssd_pnt = 1,
  reason = concat(reason, '\n', 'Add-on: General SSD (1)')
where
  server_ind = 1
  and match(product_description_edit) against ('(SSD "SOLID STATE" "SOLID_STATE") -PCIE -SATA' in boolean mode);


end //