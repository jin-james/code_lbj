# 备注：通过time和datetime模块对时间进行处理
# 1. 按日计息-固定期限还款-计算到期还款日（传入放款日期、固定期限天数）
# 2. 按月计息-每月对日还款-计算还款日期（传入放款日期、第N期）
# 3. 按月计息-每月固定日期还款-计算还款日期（传入放款日期、第N期）

import calendar
from datetime import datetime, timedelta, date


# 按月还款：ym_str-字符串格式放款日期; num-第N期; n_date-得到字符串类型还款日期
# (假设1月31号放款， 2月只有28天，28号为还款日；假设3月31号放款,4月30号为还款日; 假设4月5日放款，5月5日为还款日)
def date_calculate(ym_str, num):
    ym_date = datetime.strptime(ym_str, "%Y-%m-%d")
    s_year = ym_date.year
    s_month = ym_date.month + num  # 第N个月
    s_day = ym_date.day
    n_year = int((s_month - 1) / 12)  # 对比12个月，计算年度
    c_year = s_year + n_year
    c_month = s_month - 12 * n_year
    c_day = calendar.monthrange(c_year, c_month)[1]  # 获取某月多少天：闰年（被4整除且不被100整除、或者被400整除）2月29天,否则28天
    # print(type(days),days)                           # 第一个元素：这个月的第一天是星期几(0-6)； 第二个元素：这个月的天数
    if c_day >= s_day:  # 如果当月天数小于借款日天数，取当月天数，否则取借款日天数
        n_day = s_day
    else:
        n_day = c_day
    n_date = date(c_year, c_month, n_day).strftime('%Y-%m-%d')
    return n_date


def date_split(dt_str):
    try:
        dt_str = dt_str.split('-')
        year = int(dt_str[0])
        month = int(dt_str[1])
        day = int(dt_str[2])
    except Exception as e:
        raise ImportError("输入日期格式不对-{}".format(e))
    return year, month, day


# 传入str类型计算到期还款日（按日计息）
def month_days(dt_str, num):
    """
    :param dt_str: 日期字符串
    :param num: 天数
    :return:
    """
    year, month, day = date_split(dt_str)
    day += num
    n_date = get_date(year, month, day)
    return n_date


def get_date(year, month, day):
    """
    :param year: 年
    :param month: 月
    :param day: 日
    :return: 返回[年，月，日]
    """
    mdays = [0, 31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]
    February = 2
    while True:
        mday = mdays[month] + (month == February and isleap(year))
        if day > mday:
            day -= mday
            month += 1
        else:
            break
    if not 1 <= month <= 12:
        n, month = divmod(month, 12)
    return [year, month, day]


def date_calculate_(year, month, day):
    """
    按月为期限，默认一个月
    :param year: 年
    :param month: 月
    :param day: 日
    :return: n_date, 返回日期[年，月，日]
    """
    num = 1  # 一期（一个月）
    month = month + num
    n_year = int((month - 1) / 12)  # 12个月，如果跨年，进一
    year += n_year
    month -= 12 * n_year
    # 获取某月多少天：闰年(被4整除 and (不被100整除 or 被400整除)) 2月29天,否则28天
    # 返回这个月的天数
    n_day = get_day(year, month)
    if n_day >= day:  # 如果当月天数小于借款日天数，取当月天数，否则取借款日天数
        n_day = day
    n_date = [year, month, n_day]
    return n_date


def get_day(year, month):
    """
    :param year: 年
    :param month: 月
    :return: 返回特定年月的最大日期数, ndays:(28-31)
    """
    mdays = [0, 31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]
    February = 2
    if not 1 <= month <= 12:
        raise ImportError(month)
    ndays = mdays[month] + (month == February and isleap(year))
    return ndays


def isleap(year):
    """
    :param year: 年份
    :return: True：是闰年；False：不是闰年
    """
    return year % 4 == 0 and (year % 100 != 0 or year % 400 == 0)


# 每月固定某天还款(开始日期，期数，到期还款日)
def date_fixed(year, month, num, end_day):
    """
    :param year: 起始年
    :param month: 起始月
    :param day: 起始日(一般在0-28)？？
    :param num: 期数
    :param end_day:截止日
    :return:
    """
    s_month = month + num  # 第N个月
    n_year = int((s_month - 1) / 12)  # 对比12个月，如果跨年，进一
    year += n_year
    month = s_month - 12 * n_year
    n_date = [year, month, end_day]
    return n_date


def remove_duplication(num):
    """
    :param num: 有序数组
    :return:返回去重后的数组长度
    """
    length = len(num)
    if length <= 1:
        return num
    slow, fast = 0, 1
    for fast in range(1,length):
        if num[fast] != num[slow]:
            slow += 1
            num[slow] = num[fast]
    return slow + 1


if __name__ == '__main__':
    # n_date1 = date_calculate_(2013, 1, 29)
    # n_date2 = date_calculate_(2013, 1, 30)
    # n_date3 = date_calculate_(2000, 1, 29)
    # n_date4 = date_calculate_(2100, 1, 31)
    # n_date5 = date_calculate_(2016, 7, 31)
    # print(n_date1, n_date2, n_date3, n_date4, n_date5)
    print(month_days("2020-01-31", 30))
    print(remove_duplication([1,2,3,4,4,4,4,5,6,7,7,7,8]))
