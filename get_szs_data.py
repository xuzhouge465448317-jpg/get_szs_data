# 导入所需模块
import os  # 操作文件和目录
import json  # 处理 JSON 数据
import time  # 控制时间间隔，记录用时
import logging  # 记录日志
import requests  # 发送 HTTP 请求
import pandas as pd  # 处理 Excel 文件
import re # 正则

# 配置日志，记录到 spider.log 文件，编码为 utf-8
logging.basicConfig(
    filename="spider.log",
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    encoding="utf-8"
)


# 打印和记录日志的统一方法
def log(msg):
    print(msg)
    logging.info(msg)


# 请求相关配置
API_URL = "http://www.szse.cn/api/disc/announcement/annList"  # 年报数据接口地址
DOWNLOAD_PREFIX = "http://disc.static.szse.cn/download"  # PDF 文件下载地址前缀
SAVE_DIR = "F:\\上交所"  # 下载的 PDF 文件保存目录
EXCEL_FILE = "F:\\上交所\\年报元信息.xlsx"  # 元信息保存的 Excel 文件
DOWNLOADED_IDS_FILE = "downloaded_ids.txt"  # 已下载 annId 的记录文件

# HTTP 请求头，模拟浏览器访问


HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/138.0.0.0 Safari/537.36 Edg/138.0.0.0",
    "Accept": "application/json, text/javascript, */*; q=0.01",
    "Content-Type": "application/json",
    "Referer": "http://www.szse.cn/disclosure/listed/notice/",
    "Origin": "http://www.szse.cn",
    "X-Requested-With": "XMLHttpRequest"
}

# 创建保存 PDF 的目录，如果已经存在不会报错
os.makedirs(SAVE_DIR, exist_ok=True)


# 读取已下载过的 annId（用于断点续爬）
def load_downloaded_ids():
    if os.path.exists(DOWNLOADED_IDS_FILE):
        with open(DOWNLOADED_IDS_FILE, "r", encoding="utf-8") as f:
            return set(line.strip() for line in f if line.strip())
    return set()


def clean_filename(name: str) -> str:
    # 去掉 * ，也可以去掉其他非法字符
    return re.sub(r'[\\/:*?"<>|]', '', name)


# 保存下载成功的 annId 到本地文件
def save_downloaded_ids(ids):
    with open(DOWNLOADED_IDS_FILE, "w", encoding="utf-8") as f:
        for _id in ids:
            f.write(str(_id) + "\n")


# 将每条元信息追加保存到 Excel 表格中
def append_excel(row):
    df_new = pd.DataFrame([row], columns=["证券代码", "公司名称", "公告标题", "发布日期", "下载链接", "文件名"])
    if os.path.exists(EXCEL_FILE):
        df = pd.read_excel(EXCEL_FILE)
        df = pd.concat([df, df_new], ignore_index=True)
    else:
        df = df_new
    df.to_excel(EXCEL_FILE, index=False)


# 生成 PDF 下载链接，包含文件名参数（防止乱码）
def make_download_url(path, title):
    return f"{DOWNLOAD_PREFIX}{path}?n={requests.utils.quote(title)}.pdf"


# 下载 PDF 文件，失败时重试最多3次
def download_pdf(url, file_path, retries=3):
    for attempt in range(1, retries + 1):
        try:
            r = requests.get(url, headers=HEADERS, timeout=15)
            r.raise_for_status()
            with open(file_path, "wb") as f:
                f.write(r.content)
            return True
        except Exception as e:
            log(f"[⚠️] 第 {attempt} 次尝试失败：{e}")
            time.sleep(3)
    return False


# 替换文件名中不合法的字符
def sanitize_filename(name):
    return name.replace("/", "_").replace("\\", "_").replace("：", "_").replace(":", "_") \
        .replace("?", "").replace("*", "").replace("\"", "").replace("<", "") \
        .replace(">", "").replace("|", "")


# 主程序：从第 1
# 页开始爬取，最多 max_pages 页
def crawl_ann_reports(start_date="2024-07-06", end_date="2025-07-06", max_pages=1):
    downloaded_ids = load_downloaded_ids()  # 读取已下载记录
    total_new = 0  # 记录下载成功的数量
    total_continue = 0  # 记录跳过的数量

    for page in range(1, max_pages + 1):
        # POST 请求体
        payload = {
            "seDate": [start_date, end_date],
            "channelCode": ["listedNotice_disc"],
            "bigCategoryId": ["010303", "010301"],  # 年报类别
            "pageSize": 50,
            "pageNum": page
        }

        # 每次生成一个随机数，避免缓存
        random_value = str(time.time())
        url = f"{API_URL}?random={random_value}"
        try:
            # 发起 POST 请求
            res = requests.post(url, headers=HEADERS, data=json.dumps(payload), timeout=10)
            res.raise_for_status()
            data = res.json().get("data", [])
            log(f"✅✅✅✅✅✅✅✅✅✅✅✅✅✅✅✅请求第 {page} 页成功！✅✅✅✅✅✅✅✅✅✅✅✅✅✅✅✅")
        except Exception as e:
            log(f"❌ 请求第 {page} 页失败：{e}")
            break

        # 如果返回数据为空，则停止爬取
        if not data:
            log(f"📭 第 {page} 页无数据，结束。")
            break

        for ann in data:
            ann_id = ann.get("annId")
            title = ann.get("title", "")
            if ann_id in downloaded_ids:
                continue
            # 在循环中筛选公告时加这个判断：
            match = re.search(r"年度报告.*(更|修|补|改)", title)
            if (
                    "年度报告" not in title
                    or "摘要" in title
                    or "英文" in title
                    or match  # 年报后面有“更”“修”等修订字样
            ):
                total_continue += 1
                log(f"📄 标题含有不需字符，跳过下载 {title}，这是跳过的第{total_continue}文件")
                continue

            company = ann.get("secName", ["未知公司"])[0]
            code = ann.get("secCode", [""])[0]
            publish_date = ann.get("publishTime", "")[:10]
            attach_path = ann.get("attachPath")
            if not attach_path:
                continue

            # 构造下载地址和保存文件名
            file_url = make_download_url(attach_path, title)
            # safe_title = sanitize_filename(title)
            file_name = clean_filename(f"{title}.pdf")
            # file_name = f"{title}.pdf"
            file_path = os.path.join(SAVE_DIR, file_name)

            # 打印日志
            log(f"\n📄 正在下载 股票代码为：[{code}] 公告标题为：{title}")
            log(f"🔗 链接：{file_url}")
            log(f"📁 保存路径：{file_path}")

            # 下载文件并记录耗时
            start_time = time.time()
            success = download_pdf(file_url, file_path)
            end_time = time.time()
            cost = end_time - start_time

            if success:
                downloaded_ids.add(ann_id)  # 加入已下载集合
                append_excel([code, company, title, publish_date, file_url, file_name])  # 写入元数据
                total_new += 1
                log(f"✅ 下载成功，这是第{total_new}个文件，用时：{cost:.2f} 秒")

            else:
                log(f"❌ 下载失败，用时：{cost:.2f} 秒")

            time.sleep(2)  # 防止 IP 被封，等待 2 秒

    save_downloaded_ids(downloaded_ids)  # 保存所有已下载记录
    log(f"\n🎉 本次新增下载：{total_new} 份年报PDF数据。🎉 本次跳过下载：{total_continue} 份不需要数据，总计{total_new + total_continue}份。")


# 程序入口
if __name__ == "__main__":
    crawl_ann_reports()  # 执行主函数，开始下载年报
