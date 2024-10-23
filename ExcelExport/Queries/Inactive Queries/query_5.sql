select 
    org_name AS Organisation_Name, 
    org_file_no AS Organisation_File_Number, 
    org_cat_ent_cd AS Category_Code, 
    to_char(org_incorp_dt, 'DD-MON-YY') AS Incorporation_Date, 
    fy_end_month AS Financial_Year_End_Month, 
    org_mobile_no AS Mobile_Number, 
    org_email AS Email_Address
from vw_all_entity 
where org_name like '%LABORATORY%' and status_desc = 'ACTIVE'