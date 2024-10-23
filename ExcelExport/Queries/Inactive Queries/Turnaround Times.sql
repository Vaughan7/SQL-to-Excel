select  DSH_SERIAL_NO,
        DSD_ITM_CD,
        (SELECT ITM_CD || ' - ' || ITM_DESC FROM ROC_RF_ITEMS WHERE itm_cd = DSD_ITM_CD) ITM_DESC,
        APPL_DT,
        APL_DECISION_CD,
        apl_last_sta_cd,
        DATE_SUBMITTED,
        APL_DECISION_DT,
        (trunc(apl_decision_dt) - trunc(DATE_SUBMITTED)+1) as sub_to_app_cal_days,
        ceil((apl_decision_dt-DATE_SUBMITTED)
         - (((TRUNC(apl_decision_dt,'D')-TRUNC(DATE_SUBMITTED,'D'))/7)*2)
         + 1
         - case when (1 + TRUNC (apl_decision_dt) - TRUNC (apl_decision_dt, 'IW'))  = '6' then  1  else  0  end
         - case when (1 + TRUNC (apl_decision_dt) - TRUNC (apl_decision_dt, 'IW'))  = '7' then  2  else  0  end
         + case when (1 + TRUNC (DATE_SUBMITTED) - TRUNC (DATE_SUBMITTED, 'IW')) = '7' then  1  else  0  end)
         as sub_to_app_workdays,
        DSH_CUSER,
        DSH_CDATE,
        (trunc(DSH_CDATE) - trunc(DATE_SUBMITTED)+1) as submitted_to_ref_cal_days,
        ceil((DSH_CDATE-DATE_SUBMITTED)
         - (((TRUNC(DSH_CDATE,'D')-TRUNC(DATE_SUBMITTED,'D'))/7)*2)
         + 1
         - case when (1 + TRUNC (DSH_CDATE) - TRUNC (DSH_CDATE, 'IW'))  = '6' then  1  else  0  end
         - case when (1 + TRUNC (DSH_CDATE) - TRUNC (DSH_CDATE, 'IW'))  = '7' then  2  else  0  end
         + case when (1 + TRUNC (DATE_SUBMITTED) - TRUNC (DATE_SUBMITTED, 'IW')) = '7' then  1  else  0  end)
         as submitted_to_ref_workdays,
        APL_CUSER,
        APL_CDATE,
         APL_CDATE,
        (trunc(APL_CDATE) - trunc(DSH_CDATE)+1) as ref_to_record_cal_days,
        ceil((APL_CDATE-DSH_CDATE)
         - (((TRUNC(APL_CDATE,'D')-TRUNC(DSH_CDATE,'D'))/7)*2)
         + 1
         - case when (1 + TRUNC (APL_CDATE) - TRUNC (APL_CDATE, 'IW'))  = '6' then  1  else  0  end
         - case when (1 + TRUNC (APL_CDATE) - TRUNC (APL_CDATE, 'IW'))  = '7' then  2  else  0  end
         + case when (1 + TRUNC (DSH_CDATE) - TRUNC (DSH_CDATE, 'IW')) = '7' then  1  else  0  end)
         as ref_to_record_workdays,
        APL_LAST_STA_CD,
        APL_DECISION_USER,
        apl_decision_dt,
        (trunc(apl_decision_dt) - trunc(APL_CDATE)+1) as record_to_approval_cal_days,
        ceil((apl_decision_dt-APL_CDATE)
         - (((TRUNC(apl_decision_dt,'D')-TRUNC(APL_CDATE,'D'))/7)*2)
         + 1
         - case when (1 + TRUNC (apl_decision_dt) - TRUNC (apl_decision_dt, 'IW'))  = '6' then  1  else  0  end
         - case when (1 + TRUNC (apl_decision_dt) - TRUNC (apl_decision_dt, 'IW'))  = '7' then  2  else  0  end
         + case when (1 + TRUNC (APL_CDATE) - TRUNC (APL_CDATE, 'IW')) = '7' then  1  else  0  end)
         as record_to_approval_workdays,
        DSH_ASS_OFFICER,
        DSH_ASS_DT,
        APL_MUSER,
        APL_MDATE,
        APL_REMARKS,
        SUBMISSION_PERIOD,
        SUBMISSION_METHOD
from    (
        select   DSH_SERIAL_NO,
                 DSH_SERIAL_NO || '/' || DSD_SEQ_NO REF_NO,
                 DSD_FILE_NO FILENO,
                 DSD_ITM_CD,    DSD_QTY,    DSD_FEE,    DSD_FINE,   DSD_AMT_PAID,   DSD_PAID_F,
                 DSH_PRESENTED_BY PRESENTER_NAME,
                 --TO_CHAR (DSH_PRESENTED_DT, 'DD/MON/RRRR') DATE_SUBMITTED,
                 TO_CHAR (trunc(DSH_PRESENTED_DT, 'Day')+1, 'dd/mon/rrrr') start_week,
                 DSH_PRESENTED_DT DATE_SUBMITTED,
                 --TO_CHAR(apl_decision_dt, 'DD/MON/YYYY hh24:mi:ss') apl_decision_dt,
                 apl_decision_dt,
                 APL_DECISION_CD ,
                 --(SELECT  TO_CHAR(org_incorp_dt, 'DD/MON/yyyy') FROM    ROC_ORG_ORGANISATION WHERE   org_file_no = DSD_FILE_NO) org_incorp_dt,
                 (SELECT  org_incorp_dt FROM    ROC_ORG_ORGANISATION WHERE   org_file_no = DSD_FILE_NO) org_incorp_dt,
                 apl_appl_dt appl_dt,
                 DECODE (DSH_ASS_OFFICER, NULL, 'NOT ASSIGNED', DSH_ASS_OFFICER) DSH_ASS_OFFICER,
                 DSH_ASS_DT ,
                 apl_last_sta_cd,
                 apl_decision_user,
                 apl_cuser,         apl_cdate,
                 apl_muser,         apl_mdate,
                 dsh_cuser,         dsh_cdate,
                 APL_REMARKS,
                 TO_CHAR (DSH_PRESENTED_DT, 'RRRR-MM')submission_period,
                 nvl2(apl_job_no, 'Online','Normal') submission_method
        from    ROC_FL_DOC_SUBMIT_HDR,  ROC_FL_DOC_SUBMIT_DTLS, ROC_ORG_APPLICATIONS
        where   DSH_SERIAL_NO = DSD_DSH_SERIAL_NO
        AND     DSD_DSH_SERIAL_NO = APL_DSH_SERIAL_NO (+)
        AND     DSD_SEQ_NO = NVL (APL_DSD_SEQ_NO(+), DSD_SEQ_NO)
        --and     TRUNC (DSH_PRESENTED_DT) BETWEEN '01-Sep-2018'and '30-sep-2018'
        and     (TRUNC (DSH_cdate) BETWEEN '1-JUL-2024' and '30-SEP-2024'
        OR
        TRUNC (apl_decision_dt) BETWEEN '1-JUL-2024' and '30-SEP-2024'
        )
        --and     DSD_ITM_CD IN ('CC1', 'CC2', 'CC2B', 'CC8', 'CM5', 'CM8', 'CM2')
        and     DSH_PRESENTED_DT <= DSH_CDATE
        --and apl_last_sta_cd = 'APP'
        --and     rownum <6
)