# 深交所年报下载脚本说明

`get_szs_data.py` 用于从深交所公告接口查询指定日期范围内的上市公司公告，筛选年度报告 PDF，下载到本地目录，并把元信息追加写入 Excel。

## 功能范围

- 请求深交所公告列表接口：`/api/disc/announcement/annList`
- 筛选标题包含“年度报告”的公告
- 跳过“摘要”“英文”以及标题中疑似“更正/修订/补充/修改”的公告
- 下载公告 PDF
- 写入元信息：证券代码、公司名称、公告标题、发布日期、下载链接、文件名
- 使用 `downloaded_ids.txt` 记录已下载 `annId`，支持断点续跑
- 使用 `spider.log` 记录运行日志

## 环境依赖

建议使用 Python 3.9+。

```powershell
pip install requests pandas openpyxl
```

说明：

- `requests`：发送 HTTP 请求和下载 PDF
- `pandas`：写入 Excel 元信息
- `openpyxl`：让 pandas 能读写 `.xlsx`

## 运行前配置

脚本顶部有几个固定配置，运行前应先确认：

```python
SAVE_DIR = "F:\\上交所"
EXCEL_FILE = "F:\\上交所\\年报元信息.xlsx"
DOWNLOADED_IDS_FILE = "downloaded_ids.txt"
```

注意事项：

- 当前数据源是深交所，但默认保存目录名是 `F:\上交所`，建议改成准确且专用的目录，例如 `F:\深交所年报`。
- 运行前确保 `SAVE_DIR` 目录已经存在；当前脚本没有自动创建目录。
- 不要把 PDF、Excel、日志和 `downloaded_ids.txt` 提交到 Git。
- 如果要更换日期范围或页数，优先通过调用函数参数调整，不要盲目扩大页数。

## 运行方式

直接运行默认参数：

```powershell
cd F:\datacode\data_batch_bug
python .\get_szs_data.py
```

默认会执行：

```python
crawl_ann_reports(start_date="2025-03-13", end_date="2026-03-13", max_pages=119)
```

建议先小范围试跑：

```powershell
python -c "from get_szs_data import crawl_ann_reports; crawl_ann_reports(start_date='2025-03-13', end_date='2025-03-20', max_pages=2)"
```

确认输出、筛选规则和保存路径都正确后，再扩大日期范围和页数。

## 输出文件

| 文件 | 说明 | 是否建议提交 |
| --- | --- | --- |
| `SAVE_DIR/*.pdf` | 下载的年报 PDF | 否 |
| `EXCEL_FILE` | 年报元信息 Excel | 否 |
| `downloaded_ids.txt` | 已下载公告 ID，用于断点续跑 | 否 |
| `spider.log` | 运行日志 | 否 |

## 风险控制清单

运行前请确认以下事项，降低法律、合规、稳定性和数据质量风险：

- 只下载公开披露公告，不抓取需要登录、验证码、付费或权限控制的数据。
- 遵守目标网站的使用规则、单位合规要求和数据使用边界。
- 不要使用代理池、伪造身份、绕过封禁、绕过验证码或规避访问限制。
- 不要高频并发请求；脚本当前是串行下载，并带有随机等待，仍建议小批量运行。
- 如果出现 `403`、`429`、验证码、连接频繁失败或返回内容异常，应立即停止，降低对目标站点的压力。
- 不要把下载的数据包装成投资建议；公告内容需要二次校验，脚本不保证完整性、准确性和实时性。
- 使用公开数据做分析时，应保留来源字段和下载时间，方便追溯。
- 不要在日志或 Excel 中写入个人账号、Cookie、Token、代理地址等敏感信息。
- 脚本当前使用部分 `http://` 地址；在不可信网络环境中运行有被篡改风险，能使用 HTTPS 时应优先使用 HTTPS。
- 下载量较大时要确认磁盘空间，避免把系统盘或共享盘写满。

## 建议的安全运行流程

1. 修改 `SAVE_DIR` 和 `EXCEL_FILE` 到专用目录。
2. 手动创建保存目录。
3. 用 `max_pages=1` 或较短日期范围试跑。
4. 打开 Excel 和 PDF 抽样核对，确认筛选结果符合预期。
5. 分批运行，不要一次性扩大到过大的页数。
6. 运行过程中发现异常响应、封禁提示或日志快速增长时立即停止。
7. 运行结束后备份 Excel 元信息和 `downloaded_ids.txt`。

## 已知限制

- 脚本通过公告标题做规则筛选，可能误跳过或误下载部分公告。
- `downloaded_ids.txt` 只在本轮结束后统一保存；如果中途强制退出，本轮已下载 ID 可能没有完全落盘。
- Excel 采用追加后整表重写的方式，数据量大时会变慢。
- 当前脚本没有自动创建下载目录，也没有校验 PDF 内容是否真的有效。
- 接口字段、公告分类或站点规则变化后，脚本可能需要调整。

## 故障排查

- `FileNotFoundError`：检查 `SAVE_DIR` 是否存在。
- `ModuleNotFoundError`：执行 `pip install requests pandas openpyxl`。
- Excel 写入失败：确认 Excel 文件没有被打开占用。
- 下载重复：检查 `downloaded_ids.txt` 是否在当前工作目录下，并确认没有被删除。
- 请求失败：降低页数和频率，稍后重试；如果持续失败，不要继续高频请求。
