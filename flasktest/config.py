BASE_URL = "http://10.61.248.12:8099/mesFactoryReport"
LOGIN_URL = f"{BASE_URL}/validUserInfoByMES"
PROJECT_LIST_URL = f"{BASE_URL}/linePlanQty/queryDayPlanProject"
FPY_URL = f"{BASE_URL}/fpyReport/projectFpyReport"
NTF_DETAIL_URL = f"{BASE_URL}/fpyReport/queryNoFailDetail"
DER_DETAIL_URL = f"{BASE_URL}/fpyReport/queryFailDetail"

USER_CODE = "rti02"
PASSWORD = "888888"
LANG_CODE = "en"

NTF_GOALS = {"PCURR": 0.50, "AUD": 1.50, "ANTWBG": 2.00}
DER_GOALS = {"RQC": 1.20, "RQC2": 0.10, "MMI1": 0.35, "MMI2": 0.50, "ANTWBG": 0.70, "AUD": 0.70}
