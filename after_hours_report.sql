select
p.number 												program_number,
date_format(p.start_date,'%m-%d-%y')					start_date,
pt.description 											program_type,
pc.description 											program_category,
concat(c1.firstname, ' ', c1.lastname) 					rep,
c1.phone_mobile 										rep_phone_num,
concat(c2.firstname, ' ', c2.lastname)  				manager,
c2.phone_mobile 										manger_phone_num,
concat(c3.prefix, ' ', c3.firstname, ' ', c3.lastname) 	speaker,
c3.phone_mobile 										speaker_phone_num,
v.name 													venue_name,
p.venue_room_name 										private_room_name,

concat(v.straddr1,
case
when ifnull(v.straddr2, '') = '' then ''
else concat('\n', v.straddr2) end) 						venue_address,

v.city 													venue_city,
v.state 												venue_state,
v.phone_primary 										venue_phone,
av.name 												av_company,
av.phone_primary 										av_phone,

concat
(
case when p.av_screen = 1 then concat('screen', '\n') else '' end,
case when p.av_laptop = 1 then concat('laptop', '\n') else '' end,
case when p.av_flatscreen_tv = 1 then concat('flatscreen tv', '\n') else '' end,
case when p.av_lcd_projector = 1 then concat('lcd projector', '\n') else '' end,
case when p.av_microphone = 1 then concat('microphone', '\n') else '' end,
case when p.av_laser_pointer = 1 then concat('laser pointer', '\n') else '' end,
case when p.av_setup_delivery = 1 then concat('setup delivery', '\n') else '' end,
case when p.av_tv_cords = 1 then concat('tv cords', '\n') else '' end,
case when ifnull(p.av_other_description, '') <> '' then concat('other: ', p.av_other_description) else '' end
) 														av_order,

p.av_notes	 											av_notes

from
project pj
join program p on p.fk_project_id = pj.id
join program_type pt on p.fk_program_type_id = pt.id
join program_category pc on p.fk_program_category_id = pc.id
join vendor v on p.venue_fk_vendor_id = v.id
left join vendor av on p.av_fk_vendor_id = av.id

join client_staff_project_assignment cspa1 on p.fk_client_staff_project_assignment_id = cspa1.id
join client_staff cs1 on cspa1.fk_client_staff_id = cs1.id
join contact c1 on cs1.fk_contact_id = c1.id

left join client_staff_project_assignment cspa2 on cspa1.region = cspa2.region
  and cspa1.fk_project_id = cspa2.fk_project_id
  and cspa2.role = 'District Sales Manager'
  and cspa2.active = 1
left join client_staff cs2 on cspa2.fk_client_staff_id = cs2.id
left join contact c2 on cs2.fk_contact_id = c2.id

join registrant r on r.fk_program_id = p.id
join speaker_assignment sa on r.fk_speaker_assignment_id = sa.id
join speaker_profile sp on sa.fk_speaker_profile_id = sp.id
join speaker s on sp.fk_speaker_id = s.id
join contact c3 on s.fk_contact_id = c3.id

where
pj.code = $P{PROJECT_CODE} and
p.start_date >= $P{start_date_begin} and
p.start_date <= $P{start_date_end}