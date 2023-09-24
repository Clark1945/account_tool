import datetime
from datetime import timedelta
import pandas as pd


def test():
    pd.set_option('display.max_columns', None)
    pd.set_option('display.max_rows', None)
    today = str(datetime.date.today()).replace("-", "")
    yesterday = str(datetime.date.today() - timedelta(days=1)).replace("-", "")
    df_today = pd.read_excel("收款對帳差異報表_" + today + ".xlsx", sheet_name="交易差異檔")
    df_yesterday = pd.read_excel("收款對帳差異報表_" + yesterday + ".xlsx", sheet_name="交易差異檔")

    df_today_dist_AC_bank = df_today[df_today["差異原因"] == "[收單行無資料]、[結算時間]"]
    df_yesterday_Bank = df_yesterday[df_yesterday["差異原因"] == "[收單行無資料]、[結算時間]"]
    df_today_dist_AC = df_today[df_today["差異原因"] == "[AC無資料]、[結算時間]"]
    df_today_settlement = df_today[(df_today["差異原因"] == "[結帳日]") | (df_today["差異原因"] == "[手續費]、[結帳日]")]

    if len(df_today_dist_AC_bank) != 0: distinguish_bank_no_data(df_today_dist_AC_bank, df_yesterday)
    if len(df_today_dist_AC) != 0: distinguish_AC_no_data(df_today_dist_AC, df_yesterday_Bank)
    distinguish_settlement(df_today_settlement)

# 判斷AC無資料、結算時間
def distinguish_AC_no_data(df_today, df_yesterday):
    df_today = df_today.copy()
    df_yesterday = df_yesterday.copy()

    df_today_dist_AC = df_today.fillna("Null")
    df_yesterday_Bank = df_yesterday.fillna("Null")
    today_tx_list = sorted(df_today_dist_AC["交易序號"])
    yester_tx_list = sorted(df_yesterday_Bank["交易序號"])
    for x in range(len(today_tx_list)):
        if today_tx_list[x] in yester_tx_list:  # 如果這筆AC無資料 有存在於 昨天的 收單行無資料 且 訂單編號為 1對1對應
            if (yester_tx_list.count(today_tx_list[x]) == 1):
                continue
            else:  # 如果 今天的訂單編號在昨天存在多筆，則抓出今天的收單行交易金額，與多筆的同訂單的AC金額比對，True則不印出 ####如果同筆訂單同金額 則會出問題
                list1 = df_yesterday_Bank[df_yesterday_Bank["交易序號"] == today_tx_list[x]]["AC交易金額"].tolist()
                for x in list1: print(x)
                # if not df_today_dist_AC[df_today_dist_AC["交易序號"] == today_tx_list[x]]["收單行交易金額"].values in list1:
                #     print(df_yesterday_Bank[df_yesterday_Bank["交易序號"] == today_tx_list[x]])
        else:  # 如果這筆的AC無資料 昨天沒有出現
            print(df_today_dist_AC[df_today_dist_AC["交易序號"] == today_tx_list[x]])

        if yester_tx_list[x] in today_tx_list:  # 如果這筆收單行無資料 有存在於 昨天的 AC無資料 且 訂單編號為 1對1對應
            if (today_tx_list.count(yester_tx_list[x]) == 1):
                continue
            else:  # 如果 昨天的訂單編號在今天存在多筆，則抓出昨天的AC交易金額，與多筆的同訂單的AC金額比對，True則不印出
                list1 = df_today_dist_AC[df_today_dist_AC["交易序號"] == today_tx_list[x]]["收單行交易金額"].tolist()
                for x in list1: print(x)
                # if not df_yesterday_Bank[df_yesterday_Bank["交易序號"] == today_tx_list[x]]["AC交易金額"].values in list1:
                #     print(df_today_dist_AC[df_today_dist_AC["交易序號"] == today_tx_list[x]])
        else:  # 如果昨天訂單資料今天沒有
            print(df_yesterday_Bank[df_yesterday_Bank["交易序號"] == yester_tx_list[x]])


# 判斷收單行無資料、結算時間
def distinguish_bank_no_data(df_today, df_yesterday):
    df_today_bank = df_today.copy()
    df_today_bank["請款時間"] = pd.to_datetime(df_today_bank["請款時間"])

    df_today_dist_ctbc = df_today_bank[(df_today_bank["收單行"] == "中信")]
    df_today_dist_tspg_scrap = df_today_bank[(df_today_bank["收單行"] == "台新") & (
            (df_today_bank["AC清算名稱"] == "ＲＡＷ") | (df_today_bank["AC清算名稱"] == "ｂｌｕ　ｋｏｉ"))]
    df_today_dist_tspg_swp = df_today_bank[(df_today_bank["收單行"] == "台新") & (
            (df_today_bank["AC清算名稱"] != "ＲＡＷ") & (df_today_bank["AC清算名稱"] == "ｂｌｕ　ｋｏｉ"))]
    df_today_dist_ubot = df_today_bank[df_today_bank["收單行"] == "聯邦"]

    yesterday = datetime.date.today() - timedelta(days=1)  # 昨天的請款時間 應該為 前天
    ctbc_time = datetime.datetime(yesterday.year, yesterday.month, yesterday.day, 19, 35, 0)
    tspg_scrap_time = datetime.datetime(yesterday.year, yesterday.month, yesterday.day, 21, 5, 0)
    tspg_swp_time = datetime.datetime(yesterday.year, yesterday.month, yesterday.day, 22, 5, 0)
    ubot_time = datetime.datetime(yesterday.year, yesterday.month, yesterday.day, 21, 30, 0)

    df_today_dist_ctbc = df_today_dist_ctbc[df_today_dist_ctbc["請款時間"] <= ctbc_time]  # 中信昨晚八點以前的交易
    if len(df_today_dist_ctbc) != 0: print(df_today_dist_ctbc)
    df_today_dist_tspg_scrap = df_today_dist_tspg_scrap[
        df_today_dist_tspg_scrap["請款時間"] <= tspg_scrap_time]  # 爬帳昨晚九點以前的交易
    if len(df_today_dist_tspg_scrap) != 0: print(df_today_dist_tspg_scrap)
    df_today_dist_tspg_swp = df_today_dist_tspg_swp[
        df_today_dist_tspg_swp["請款時間"] <= tspg_swp_time]  # SWP 昨晚十點以前的交易
    if len(df_today_dist_tspg_swp) != 0: print(df_today_dist_tspg_swp)
    df_today_dist_ubot = df_today_dist_ubot[df_today_dist_ubot["請款時間"] <= ubot_time]  # 聯邦九點半以前的交易
    if len(df_today_dist_ubot) != 0: print(df_today_dist_ubot)


# 判斷結帳日
def distinguish_settlement(df_today_settlement):
    yesterday = datetime.date.today() - timedelta(days=1)
    yesterday_tx1 = datetime.date.today() - timedelta(days=2)
    yesterday_tx2 = datetime.date.today() - timedelta(days=3)
    settle_time = str(datetime.date(yesterday.year, yesterday.month, yesterday.day)).replace("-", "")
    settle_time_tx1 = str(datetime.date(yesterday_tx1.year, yesterday_tx1.month, yesterday_tx1.day)).replace("-", "")
    settle_time_tx2 = str(datetime.date(yesterday_tx2.year, yesterday_tx2.month, yesterday_tx2.day)).replace("-", "")

    # 周一結帳日 理應出現在 周五與周六
    df_settlement_abnormal = df_today_settlement[(df_today_settlement["收單行結帳日期"] == settle_time) & (
            (df_today_settlement["AC結帳日期"] != settle_time_tx1) | (
            df_today_settlement["AC結帳日期"] != settle_time_tx2))]
    if len(df_settlement_abnormal) != 0: print(df_settlement_abnormal)
    df_settlement_exception = df_today_settlement[(df_today_settlement["收單行結帳日期"] != settle_time)]
    if len(df_settlement_exception) != 0: print(df_settlement_exception)
