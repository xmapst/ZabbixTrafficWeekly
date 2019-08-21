#!/usr/bin/python3

from zabbix_api import ZabbixAPI
import math
import datetime
import time
import xlwt
import os
import threading

ZABBIX_SERVER = 'http://localhost/zabbix'
TIMEOUT = 60
USER = "Admin"
PASSWD = "zabbix"

EXCEL_PATH = "E:\python\\"
EXCEL_NAME = "network_montor.xls"

def cover_excel(msg,start_time):
    # 写入到excel表格
    ws = wb.add_sheet(start_time, cell_overwrite_ok=True)
    count = len(msg)
    x = msg
    # 表头
    title = ["时间", "组名", "主机IP", "别名", "网卡名", "进入最小流量", "进入平均流量", "进入最大流量", "流出最小流量", "流出平均流量", "流出最大流量"]
    x.insert(0, title)

    # 表内容
    ## 列
    for j in range(0, 11):
        ## 行
        for i in range(0, count):
            if i == 0:
                value = x[0]
            else:
                value = x[i]
            ## 转换类型
            if isinstance(value[j], int) or isinstance(value[j], float):
                ws.write(i, j, value[j])
            else:
                ws.write(i, j, value[j])

def clockTotime(value):
    # 时间戳转换为时间
    clock = time.localtime(int(value))
    format = "%Y-%m-%d %H:%M"
    return time.strftime(format, clock)

def timeToclock(value):
    # 时间转换为时间戳
    colock = time.strptime(value, '%Y-%m-%d')
    return int(time.mktime(colock))

def byteFormat(size):
    # 字节单位转换
    if size > math.pow(1024, 3):
        return '%.2f G' % (size / math.pow(1024, 3))
    if size > math.pow(1024, 2):
        return '%.2f M' % (size / math.pow(1024, 2))
    if size > 1024:
        return '%.2f K' % (size / math.pow(1024, 1))
    return size

def chunkIt(seq, num):
  avg = len(seq) / float(num)
  out = []
  last = 0.0

  while last < len(seq):
    out.append(seq[int(last):int(last + avg)])
    last += avg

  return out

def get_data(search_key, start_time, end_time):
    # 根据查询key，开始及结束时间查询所有主机的item的趋势值中最小、平均、最大值
    time_from = timeToclock(start_time)
    time_till = timeToclock(end_time)

    units_list = ["B", "vps", "bps", "sps"]
    msg = []

    def getValueList(group_name, host_id, host_ip, host_name):
        lists = []
        # 获取item列表
        try:
            item_list = zapi.item.get({"output": ["itemid", "name", "key_", "value_type", "units"], "filter": {"hostid": host_id}, "search": {"name": search_key}})
            for item in item_list:
                # dervice 为网卡名
                dervice = item["key_"].split("[")[1].split("]")[0]
                # 返回值的类型以及单位
                value_type = item["value_type"]
                units = item["units"]
                # 获取item的值
                results = zapi.trend.get({"output": "extend", "itemids": item["itemid"], "time_from": time_from, "time_till": time_till})

                for i in results:
                    clock = clockTotime(i["clock"])
                    if start_time in clock:
                        # 检测是 item 类型是否是整数或者浮点数，不是则直接返回-1
                        if value_type == "3" or value_type == "0":
                            value_min, value_avg, value_max = i["value_min"], i["value_avg"], i["value_max"]
                            if value_type == "3":
                                value_min = int(value_min)
                                value_avg = int(value_avg)
                                value_max = int(value_max)
                                if units in units_list:
                                    value_min = byteFormat(value_min)
                                    value_avg = byteFormat(value_avg)
                                    value_max = byteFormat(value_max)

                            value_min = str(value_min) + units
                            value_avg = str(value_avg) + units
                            value_max = str(value_max) + units
                        else:
                            value_min = "-1"
                            value_avg = "-1"
                            value_max = "-1"
                        # 拼接成列表
                        lists = [clock, group_name, host_ip, host_name, dervice, value_min, value_avg, value_max]
                        #print("Msg：%s" % lists)
                    if lists:
                        msg.append(lists)

        except Exception as err:
            print(err)

    def threads(group_name, host_list):
        threads = []
        try:
            for host in host_list:
                host_id = host["hostid"]
                host_name = host["name"]
                host_ip = host['host']
                t = threading.Thread(target=getValueList, args=(group_name, host_id, host_ip, host_name))
                threads.append(t)

            for thread in threads:
                thread.setDaemon(True)
                thread.start()

            for stop in range(len(host_list)):
                threads[stop].join()

        except Exception as err:
            print(err)

    # 获取组列表
    group_list = zapi.hostgroup.get({"output": ["groupid", "name"]})
    for group in group_list:
        group_name = group["name"]
        host_list = zapi.host.get({"output": ["hostid", "name", "host"], "filter": {"groupids": group["groupid"]}})
        if host_list:
            split = 100
            if len(host_list) >= split:
                if (len(host_list) % split) == 0:
                    num = len(host_list) / split
                else:
                    num = len(host_list) / split + 1
                for lists in chunkIt(host_list, num):
                    threads(group_name, lists)
            else:
                threads(group_name, host_list)

        if msg:
            return msg

def main():
    # 查询前7天的数据，每天每小时的最小，平均，最大值
    # 查询的key
    in_key = "Incoming network traffic"
    out_key = "Outgoing network traffic"
    for i in range(7, 0, -1):
        start_time = ((datetime.datetime.now() - datetime.timedelta(days=i))).strftime("%Y-%m-%d")
        end_time = ((datetime.datetime.now() - datetime.timedelta(days=i - 1))).strftime("%Y-%m-%d")
        in_results = get_data(in_key, start_time, end_time)
        out_results = get_data(out_key, start_time, end_time)
        if in_results:
            for a in in_results:
                if out_key:
                    for b in out_results:
                        if len(a) == 11:
                            continue
                        if a[0] == b[0] and a[1] == b[1] and a[2] == b[2] and a[3] == b[3] and a[4] == b[4]:
                            a.append(b[5])
                            a.append(b[6])
                            a.append(b[7])
                            print(a)

            msg = in_results
            cover_excel(msg, start_time)

if __name__ == "__main__":
    # 登录zabbix
    zapi = ZabbixAPI(server="{}/api_jsonrpc.php".format(ZABBIX_SERVER), timeout=TIMEOUT)
    zapi.login(USER, PASSWD)

    if os.path.exists(EXCEL_PATH) is False:
        os.mkdir(EXCEL_PATH)

    EXCEL = EXCEL_PATH + EXCEL_NAME

    print("程序开始执行》》》")
    print("请耐心等待》》》》》》")
    wb = xlwt.Workbook()
    main()
    wb.save(EXCEL)
    print("结果保存在: 》》》》》》》》》\n", EXCEL)
