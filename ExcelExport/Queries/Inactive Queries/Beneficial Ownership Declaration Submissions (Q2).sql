select 	HDR.BO1H_APPLICATION_NUMBER,
	NVL (HDR.BO1H_DECISION, 'PEN') BO1H_DECISION, 
	HDR.BO1H_DECISION_USER,
	HDR.BO1H_DECISION_DATE,
	HDR.BO1H_CDATE, 
	DTL1.BO1D_PURPOSE,
    CASE DTL1.BO1D_PURPOSE
    WHEN 'CAMD' THEN 'Enttity Amendments'
    WHEN 'RCY' THEN 'Company Registration'
    WHEN 'RCC' THEN 'Close Corporation Registration'
    WHEN 'SAD' THEN 'Annual Duty Submission'
    WHEN 'OTH' THEN 'Other'
    ELSE 'Not Specified' END AS "BO Reason",
 	DTL1.BO1D_REGISTRATION_NUMBER,
	DTL1.BO1D_COMPANY_TYPE,
	DTL2.BO2D_PURPOSE, 
	DTL2.BO2D_ID_NUMBER,
	NVL (HDR.BO1H_DECISION, 'PEN') DUP_BO1H_DECISION,
	NVL (HDR.BO1H_DECISION, 'PEN') AS DUP2_BO1H_DECISION,
	NVL (HDR.BO1H_DECISION, 'PEN') AS DUP3_BO1H_DECISION,
	DTL2.BO2D_PURPOSE AS DUP_BO2D_PURPOSE,
CASE DTL1.BO1D_PURPOSE
    WHEN 'CAMD' THEN 'Enttity Amendments'
    WHEN 'RCY' THEN 'Company Registration'
    WHEN 'RCC' THEN 'Close Corporation Registration'
    WHEN 'SAD' THEN 'Annual Duty Submission'
    WHEN 'OTH' THEN 'Other'
    ELSE 'Not Specified' END AS "DUP BO Reason"
FROM BO1_HEADER HDR JOIN BO1_DETAILS DTL1
    ON HDR.BO1H_APPLICATION_NUMBER = DTL1.BO1D_APPLICATION_NUMBER
    JOIN BO2_DETAILS DTL2
    ON HDR.BO1H_APPLICATION_NUMBER = DTL2.BO2D_APPLICATION_NUMBER
    WHERE TRUNC (BO1H_CDATE) between '01-JUL-2024' and '30-SEP-2024'