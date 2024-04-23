# 说明

常用存货编码：

| 存货编码  | 存货名称               | 报货日期  | 备注   |
| --------- | ---------------------- | --------- | ------ |
| 040000045 | 柠檬                   |           |        |
| 040000006 | 橙子                   |           |        |
| 040001022 | 调味糖浆               |           |        |
| 120001482 | 糯香观音组合套装       | 2024/2/20 | 420杯  |
| 120001511 | 碧玉桃花组合套装       | 2024/3/7  | 420杯  |
| 120001535 | 清风茉白升级宣传物料包 | 2024/3/13 |        |
| 040001536 | 2024版清风茉白-A套餐   | 2024/3/15 | (新)   |
| 040001537 | 2024版清风茉白-B套餐   |           | (旧) |
| 120001549 | 牛油果宣传物料包       | 2024/3/23 |        |
| 040000807 | 冷冻牛油果泥           | 2024/3/23 |        |
| 060000019 | PLA粗吸管              |           |        |
| 060000020 | PLA细吸管              |           |        |
| 060000889 | PLA粗吸管（黑）        |           |        |
| 040001410 | 龙麟香组合套装         | 2024/1/6  |        |
| 040000983 | 速冻桑葚浆             |           |        |
| 120001554 | 超仙黑桑葚宣传物料包   | 2024/4/9  | 4800套 |
| 040001555 | 2024版桑葚包材套装     | 2024/4/9  | 4500套 |



常用sql

```sql
-- 创建一个临时表，用于存储查询的开始和结束日期
WITH variables AS (
  SELECT '20240401' AS start_date, '20240407' AS end_date
)
-- 查询指定日期范围内的门店账单数据
SELECT 
  -- 将开始和结束日期拼接成一个字符串，表示查询的时间段
  CONCAT(variables.start_date, '~', variables.end_date) AS 时段,
  business_date AS '日期', -- 账单日期
  stat_shop_id AS '门店编码', -- 门店编码
  SUM(order_count) AS '账单数', -- 账单数量
  SUM(order_count_last_year) AS '同期账单数', -- 同期账单数量
  ROUND(SUM(total_amount), 2) AS '流水金额', -- 流水金额，保留两位小数
  ROUND(SUM(pay_amount), 2) AS '实收金额', -- 实收金额，保留两位小数
  ROUND(SUM(total_amount_last_year), 2) AS '同期流水', -- 同期流水金额，保留两位小数
  ROUND(SUM(pay_amount_last_year), 2) AS '同期实收' -- 同期实收金额，保留两位小数
FROM ads_dbs_trade_shop_di, variables -- 从ads_dbs_trade_shop_di表和variables临时表中查询数据
WHERE business_date BETWEEN variables.start_date AND variables.end_date -- 筛选指定日期范围内的数据
GROUP BY 时段, 日期, 门店编码 -- 按时间段、日期和门店编码进行分组
ORDER BY 日期 ASC; -- 按日期升序排列结果
```

