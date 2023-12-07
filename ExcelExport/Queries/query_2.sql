SELECT 
	CASE 
		WHEN GROUPING (Divisions) = 1 THEN 'Grand Total'
			ELSE (Divisions)      
		END AS Division,

	SUM(CASE WHEN org_cat_ent_cd = '21' THEN Total ELSE 0 END) AS "21",
	SUM(CASE WHEN org_cat_ent_cd = 'CY' THEN Total ELSE 0 END) AS "CY",
	SUM(CASE WHEN org_cat_ent_cd = 'CC' THEN Total ELSE 0 END) AS "CC",
	SUM(CASE WHEN org_cat_ent_cd = 'FOR' THEN Total ELSE 0 END) AS "FOR",
	SUM(CASE WHEN org_cat_ent_cd = 'DN' THEN Total ELSE 0 END) AS "DN",
	SUM(Total) AS Grand_Total

FROM 
	(SELECT
		major_div_code AS Divisions,
		org_cat_ent_cd,
		COUNT(1) AS Total
	FROM icrs_interface.vw_all_entity
	GROUP BY major_div_code, org_cat_ent_cd) sub

GROUP BY ROLLUP (Divisions)
ORDER BY Divisions;
