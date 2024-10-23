SELECT 	PER_SEX,
	trunc(months_between(sysdate,PER_DOB)/12) PER_AGE 
FROM ICRS_INTERFACE.VW_ALL_ENTITY_PERSON
WHERE OWNERSHIP IN ('DIRECTOR','SHAREHOLDERS','SIGNATORIES TO ARTICLES','SECRETARY')
AND PER_SEX IS NOT NULL 
AND PER_DOB IS NOT NULL
AND PER_DOB < SYSDATE
AND ORG_FILE_NO IN (
select * 
from temp_table
)
 