# 西友 DSS 共通框架

个别处理只负责业务。触发、读主数据、日志、邮件一律走 `SEIYUCOM*`。

## 项目分工

| 项目 | 角色 | 能改吗 |
|---|---|---|
| `SEIYUCOM000X` | 入口。`01frameworkBatch` 由调度触发器、JP1、ScriptRunner 调用 | 只用，不改 |
| `SEIYUCOM001X` | `01getInterfaceInfo` / `02outputLog` / `03sendMail` / 调个别脚本 | 全项目共用，禁止修改 |
| `SEIYUCOM002X` | ADLS 路径切割库 | 可用 |
| `SEIYUCOM003X` | 备份清理（每天 0 点） | 了解即可 |
| HOUSEKEEP | DSS 本地备份清理 | 非推荐；备份放 Blob `AZURE_STR_DSS` |
| 接口 ID 项目 | 个别处理，被 001X 调用 | 这里写业务 |

调用链：触发器 → `01frameworkBatch` → 读 `if_master` / `param_master` → 调个别服务 → 按戻り値写日志 / 发邮件。

JP1 / ScriptRunner：共通模板 `C:\Temp\scriptrunner.xml`，第一参数传 `if_id`。日志进 My Log，和触发器启动一样。

## 四条收口（Day2）

框架先读 `if_master`、写开始日志、调个别处理，再按结束状态分四路。

| 分支 | 个别处理怎么结束 | 框架做什么 |
|---|---|---|
| 1 正常 | End 戻り値 **0**（默认） | 正常日志后结束 |
| 2 警告 | 戻り値 **1–254**（演示常用 0 件数据） | 警告邮件 + 警告日志。例：文件只有表头 |
| 3 可控异常 | 戻り値 **255**，脚本已设 `out_*` | 异常邮件 + 异常日志 |
| 4 未捕获 | Exception / 未 catch | catch 后异常邮件。戻り値实际为 **-1** |

戻り値设在 **End 组件**的属性检查器，不是脚本属性随便写一个数。

警告/异常用「警告变量设定」「异常变量设定」，把 catch 内容塞进：

- `out_component_name` / `out_component_type`
- `out_message_category` / `out_message_code` / `out_message_level`
- `out_error_type` / `out_error_message`

邮件后半段就是这些字段，例如：

```
FILE0001E|java.io.FileNotFoundException|/data/DSSROOT/DEMO/INPUT/OrderList.csv|CSV読取
```

发信：`no-reply-HR-EAI@seiyu.com`。

个别处理必须：登记 `if_master`、设 End 戻り値。可选：`param_master` 入参、输出 `out_*`。

Day2 改写口径：正常 = 跑完；警告 = 0 件；异常 = 读或写失败。

## if_master

库：`DSSSEIYU`。改前按リリース手順书 3.1 备份。发版常用同 `if_id` **DELETE 再 INSERT**。

必填列：

| 列 | 含义 |
|---|---|
| `if_id` | 接口 ID = 项目名 |
| `if_name` | 显示名 |
| `service_name` | 个别项目登记后的服务名（通常 = 项目名） |
| `script_name` | 个别脚本名，如 `01OutputData` |
| `from` / `to` / `cc` / `bcc` | 邮件。**to 必须含** `isd_support_desk@seiyu.com` |
| `mail_title_header` | 邮件标题头 |
| `tier` | 重要度 |

## param_master

`IF_ID` + `param_no` + `param_value`。`01getInterfaceInfo` 读入后塞到 `param1`–`param20`。

个别脚本用这些变量表达环境差异，**不要**把路径/文件名写死在图标属性里（除非台账规定该 IF 无参数）。
