SELECT ORG_FILE_NO, ORG_CAT_ENT_CD, MAJOR_DIV_CODE, ORG_LAST_STA_CD, ORG_INCORP_DT
from VW_ALL_ENTITY
WHERE TRUNC (ORG_INCORP_DT) <= '30-SEP-2024'
AND org_last_sta_cd like 'A%'
order by ORG_INCORP_DT desc