SELECT VOTE_ITEM,
         SUM (RPD_FEE) FEE,
         SUM (RPD_EXAM_FEE) EXAM_FEE,
         SUM (RPD_FINE) ITEM_FINE,
         SUM (RPD_AMT_PAID) AMT_PAID
    FROM (  SELECT RPH_RECEIPT_NO,
                   CASE
                      WHEN NVL (RPD_TO_DT, SYSDATE) < '31-DEC-2022'
                      THEN
                         DECODE (RPD_ITM_CD,
                                 'CM23', '1050/030',
                                 'CM23B', '1050/030',
                                 'CC7', '1050/030',
                                 NVL (DTL.RPD_REV_ITEM_CD, 'ABCD'))
                      ELSE
                         NVL (DTL.RPD_REV_ITEM_CD, 'ABCD')
                   END
                      VOTE_ITEM,
                   SUM (RPD_FEE) RPD_FEE,
                   SUM (RPD_EXAM_FEE) RPD_EXAM_FEE,
                   SUM (RPD_FINE) RPD_FINE,
                   SUM (RPD_AMT_PAID) RPD_AMT_PAID
              FROM ROC_CO_RECEIPT_DTL DTL, ROC_CO_RECEIPT_HDR HDR
             WHERE HDR.RPH_RECEIPT_NO = DTL.RPD_RPH_RECEIPT_NO
                   AND TRUNC (HDR.RPH_PAYMENT_DT) BETWEEN '01-NOV-2023'
                                                      AND  SYSDATE
          GROUP BY RPH_RECEIPT_NO,
                   CASE
                      WHEN NVL (RPD_TO_DT, SYSDATE) < '31-DEC-2022'
                      THEN
                         DECODE (RPD_ITM_CD,
                                 'CM23', '1050/030',
                                 'CM23B', '1050/030',
                                 'CC7', '1050/030',
                                 NVL (DTL.RPD_REV_ITEM_CD, 'ABCD'))
                      ELSE
                         NVL (DTL.RPD_REV_ITEM_CD, 'ABCD')
                   END)
GROUP BY VOTE_ITEM
ORDER BY VOTE_ITEM