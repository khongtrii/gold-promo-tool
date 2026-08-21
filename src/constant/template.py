from typing import List

column_check_oa: List[str] = [
    "GOLD CODE", "LV", "LU", "SITE", "SUPPLIER", 
    "ADDRESS CHAIN", "CONTRACT", "START DATE", "END DATE"
]

column_promotion_plan: List[str] = [
    "SO", "DESCRIPTION", "THEME", "SITE GROUP", "ACTION", "SALE START", "SALE END",
    "GLOBAL PERIOD START", "GLOBAL PERIOD END", "SHOP ACTIVATION",
    "COMMITMENT DEADLINE", "COMMITMENT CLOSING", "ORDER WH START", "ORDER WH END"
]

column_update_so: List[str] = [
    "ACTION", "SO", "GOLD CODE", "LV", "LU"
]

column_missing_ou: List[str] = [
    "SO", "SITE", "GOLD CODE", "LV", "LU",
    "SUPPLIER", "ADDRESS CHAIN", "CONTRACT", "PRIORITY TYPE", "SO TYPE"
]

column_so_calendar: List[str] = [
    "SO", "SITE", "SUPPLIER", "CONTRACT", "ADDRESS CHAIN",
    "GOLD CODE", "LV", "LU", "ORDER DATE", 
    "DELIVERY DATE", "PCT WEIGHT", "CALENDAR TYPE", "MESSAGE"
]

column_purchase: List[str] = [
    "GOLD CODE", "LV", "APPLY UNIT", "PURCHASE PRICE", "SITE", 
    "START DATE", "END DATE", "CONTRACT", "VAT", "SUPPLIER", "ACTION"
]

column_po_commitment: List[str] = [
    "PROMOTION PLAN", "SO", "GOLD CODE", "LV", "LU",
    "SITE", "QUANTITY", "SO TYPE"#, "SUPPLIER"
]

column_supplier_schedule: List[str] = [
    "SO", "SITE", "SUPPLIER", "ADDRESS CHAIN", "CONTRACT",
    "SERVICE CONTRACT", "ORDER DATE", "ORDER TIME",
    "DELIVERY DATE", "DELIVERY TIME", "SCHEDULE TYPE",
    "DELIVERY NO", "MESSAGE"
]

column_ag: List[str] = [
    "ACTION", "DEPARTMENT", "SITE", "SUPPLIER", "CONTRACT", "AG NO", "AG CODE", "AG DESCRIPTION",
    "AG START DATE", "AG END DATE", "GOLD CODE", "LV", "ARTICLE START DATE", "ARTICLE END DATE", "LEVEL OF VALUE", "ERROR"
]

column_dc: List[str] = [
    "ACTION", "SITE", "SUPPLIER", "CONTRACT", "AG CODE", "AG DESCRIPTION", "DISCOUNT CODE", "DISCOUNT TYPE",
    "DOCUMENT TYPE", "APPLICATION", "DISCOUNT SUB_TYPE", "CO", "AO", "START DATE", "END DATE", "DISCOUNT UNIT",
    "VALUE ON INVOICE", "APPLY UNIT", "GOLD CODE", "LV", "CONDITION QUANTITY", "CONDITION UNIT", 
    "FREE QUANTITY", "FREE UNIT", "DETAIL START DATE", "DETAIL END DATE", "ERROR"
]

column_de: List[str] = [
    "ACTION", "SITE", "AG CODE", "DISCOUNT CODE", "GOLD CODE", "LV",
    "VALUE ON INVOICE", "START DATE", "END DATE", "COMMENT", "ERROR"
]

report: List[str] = [
    'STT', 'ACTION', 'DEPT', 'SITE', 'SUPPLIER_CODE', 'COMERCIAL_CONTRACT', 
    'AGNO1', 'AG_CODE', 'AG_DESC', 'AG_START_DATE', 'AG_END_DATE', 'ARTICLE_CODE', 
    'LV', 'ARTICLE_START_DATE', 'ARTICLE_END_DATE', 'LEVEL_OF_VALUE', 
    'STT_COMERCIAL_CONTRACT', 'STT_ARTICLE_GROUP', 'FILEID', 'FILENM', 
    'USERNM', 'DATECRE', 'ACTI', 'ERRORMESS'
]

column_add_attribute_marketing: List[str] = [
    "ACTION", "SO PLAN", "GOLD CODE", "LV", "LU", "MEDIUM", "PHOTO"
]

column_sale: List[str] = [
    "ACTION", "GOLD CODE", "SV", "SALE PRICE", "PRICELIST", "START DATE", "END DATE", "VAT", "REASON"
]

column_attr: List[str] = [
    'ACTION','GOLD CODE','SV','CLASS','CODE','ALPHANUM','NUM_VALUE','DATE','TIME','START DATE','END DATE','NOTE COUNT'
]
