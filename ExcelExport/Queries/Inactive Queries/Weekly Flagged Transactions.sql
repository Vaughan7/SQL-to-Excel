SELECT TRUNC (DSD_CDATE) AS "Date of transaction",
         DSD_ITM_CD AS "Form name",
         ORG_NAME AS "Entity name",
         DSD_FILE_NO AS "Entity registration number",
         NVL (DSD_PAID_F, 'N') AS "Flagged",
         DSD_CUSER AS "Staff who initiated",
         DSD_MUSER AS "Staff member who modified", 
         DSD_PAY_FLAG_USER AS "Staff member who flagged",
         DSD_PAY_FLAG_DATE AS "Flagging date",
         DSD_PAY_FLAG_REASON AS "Reason for flagging",
         nvl(DSD_FEE,0) AS "Amount flagged",  
         DSD_MDATE AS "Date of modification",
         DSH_FTP_FILING_TYPE,
         FTP_FILING_DESC,
         DSD_DSH_SERIAL_NO,
         DSD_SEQ_NO,
         dsd_itm_cd, 
         ITM_DESC,
         --EXTRACT(Year FROM DSD_CDATE) AS "Year (Date of Transactions)",
         --TO_CHAR(DSD_CDATE, 'Month') AS "Month (Date of Transaction)",
         --TO_CHAR(DSD_CDATE, 'Day') AS "Day (Date of Transaction)",
         TO_CHAR(DSD_PAY_FLAG_DATE, 'DD Month, DAY')
    FROM ROC_FL_DOC_SUBMIT_HDR,
         ROC_FL_DOC_SUBMIT_DTLS,
         ROC_ORG_ORGANISATION,
         ROC_RF_ITEMS,
         ROC_RF_FILING_TYPES
   WHERE     DSH_SERIAL_NO = DSD_DSH_SERIAL_NO
         AND ITM_CD = DSD_ITM_CD
         AND FTP_FILING_TYPE = DSH_FTP_FILING_TYPE
         AND DSD_FILE_NO = ORG_FILE_NO
         AND NVL (DSD_PAID_F, 'N') = 'Y'
         AND TRUNC (DSH_CDATE) between '01-APRIL-2024' and SYSDATE
         AND (DSD_PAY_FLAG_DATE >= trunc( trunc( SYSDATE, 'IW' )  - 1, 'IW') 
                and DSD_PAY_FLAG_DATE < trunc( SYSDATE, 'IW' ))
ORDER BY 1, 2