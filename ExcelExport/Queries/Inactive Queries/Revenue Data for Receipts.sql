SELECT RPH_RECEIPT_NO RECEIPT_NUMBER,
         --RPD_SEQ_NO,
         RPD_FEE AMOUNT,
         RPD_FINE , 
         RPD_AMT_PAID     ,
         RPD_ITM_CD ITEM_CODE,
         rpd_period,
         RPD_LINK_KEY1_REF REFERENCE_NUM,
         RPD_LINK_KEY2_REF SEQ,
         RPD_LINK_KEY3_REF REFERENCE3,
         rph_cuser     ISSUED_BY,
         RPH_PAYMENT_DT RECEIPT_DATE,
         rph_cdate date_captured, 
         rph_rcpt_status STATUS,
         RPD_ENT_CD ENTITY_TYPE,
         RPD_FILE_NO FILE_NUMBER,
         (SELECT ORG_NAME FROM ROC_ORG_ORGANISATION WHERE ORG_FILE_NO = RPD_FILE_NO) ORG_NAME
    FROM ROC_CO_RECEIPT_DTL, ROC_CO_RECEIPT_HDR
   WHERE     rpd_rph_receipt_no = rph_receipt_no
         AND rph_rcpt_status = 'PAY'
         --and rph_cuser in ('ZIPPORA',   'MERYAM')
         and (
         (rph_cdate >= trunc( trunc( SYSDATE, 'IW' )  - 1, 'IW') 
                and RPH_CDATE < trunc( SYSDATE, 'IW' ))
         or 
         (Rph_payment_dt >= trunc( trunc( SYSDATE, 'IW' )  - 1, 'IW') 
                and Rph_payment_dt < trunc( SYSDATE, 'IW' ))
         ) 
ORDER BY RPH_PAYMENT_DT, rph_cuser
