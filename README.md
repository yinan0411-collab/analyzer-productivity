# LAX2 出库产能与订单结构分析（初版）

## 当前功能

- 上传“考核单”和“拣货结果”Excel；两类文件均支持一次多选一周数据。
- 通过订单号关联生产需求与实际拣货。
- 以“生产结束时间”判断订单应完成日期和时点。
- 区分当天做当天、当天做未来、逾期补做。
- 订单分类：
  - `SPB名称`含“2B” → 2B；
  - 非2B且单订单件数超过可调阈值 → 超大异常单；
  - 其余 → 2C。
- 储位分类：
  - 默认 A01–A36 且 L3及以上 → 高层拣选；
  - 其余全部 → 地面拣选。
- 交叉分析 2B/2C 与高层/地面结构。
- 按任务领取时间—拣货完成时间计算任务时长，并合并员工重叠区间，生成初步操作人效。
- 下载员工、订单、任务和关联明细 CSV。

## 本地运行

1. 安装 Python 3.11 或 3.12。
2. 解压项目。
3. 双击 `run_app.bat`。
4. 浏览器打开后上传两张 Excel。

也可在命令行运行：

```bash
pip install -r requirements.txt
streamlit run app.py
```

## 发布到 Streamlit Community Cloud

1. 在 GitHub 新建 repository。
2. 上传本文件夹内的全部文件，包括 `.streamlit/config.toml`。
3. 在 Streamlit Community Cloud 选择该 repository。
4. Main file path 填写 `app.py`。
5. Deploy。

## 下一版预留

- 人员名单、人员类型、技能矩阵。
- 实际出勤与排班工时。
- 操作人效与排班人效对比。
- 复核、Babylist打包、Mix打包、2B大宗打包、2B拣货、机区拣货、退供、道口、5S。
- 一周订单量—人数—人效关联与下周人数预测。
