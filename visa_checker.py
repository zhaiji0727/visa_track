import logging
import traceback
import time
import ddddocr
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# 配置日志
logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s: %(message)s")

# -------------------------
# Configuration
BROWSER = ""  # 浏览器类型 eg. "chrome", "firefox", "safari", "edge"
webdriver_path = None  # WebDriver 路径，Safari 不需要指定(None) eg. "/path/to/chromedriver"

APP_REF = ""  # Application Reference Number eg. "HUN/PEK/xxxxxx/xxxx/xx"
LAST_NAME = ""  # Last Name eg. "ZHANG"
MAX_ATTEMPTS = 5

SEND_EMAIL = True  # 是否发送邮件通知
SENDER = "" # 发送者名称 eg. "Visa Checker"
SENDER_EMAIL = "" # 发送者邮箱
SENDER_PASSWORD = "" # 发送者邮箱密码
SMTP_SERVER = ""   # SMTP 服务器 eg. "smtp.example.com"
SMTP_PORT = 587
RECEIVER_EMAIL = "" # 接收者邮箱
# -------------------------
logging.info(f"Application Reference Number: {APP_REF}")
logging.info(f"Last Name: {LAST_NAME}")

ocr = ddddocr.DdddOcr(show_ad=False)
driver = None

try:
    # 根据 BROWSER 实例化对应的 WebDriver
    if BROWSER.lower() == "chrome":
        if webdriver_path:
            # 动态导入 Chrome Service 并构造 driver
            from selenium.webdriver.chrome.service import Service as ChromeService

            service = ChromeService(executable_path=webdriver_path)
            driver = webdriver.Chrome(service=service)
        else:
            driver = webdriver.Chrome()

    elif BROWSER.lower() == "safari":
        driver = webdriver.Safari()

    elif BROWSER.lower() == "firefox":
        if webdriver_path:
            # 动态导入 Firefox Service 并构造 driver
            from selenium.webdriver.firefox.service import Service as FirefoxService

            service = FirefoxService(executable_path=webdriver_path)
            driver = webdriver.Firefox(service=service)
        else:
            driver = webdriver.Firefox()

    elif BROWSER.lower() == "edge":
        if webdriver_path:
            # 动态导入 Edge Service 并构造 driver
            from selenium.webdriver.edge.service import Service as EdgeService

            service = EdgeService(executable_path=webdriver_path)
            driver = webdriver.Edge(service=service)
        else:
            driver = webdriver.Edge()

    else:
        raise ValueError(f"Unsupported browser: {BROWSER}")

    driver.get("https://visa.vfsglobal.com/chn/zh/hun/apply-visa")

    WebDriverWait(driver, 30).until(
        EC.presence_of_element_located((By.ID, "viewmore5"))
    )
    # 展开并点击追踪链接
    driver.execute_script("document.getElementById('viewmore5').click();")
    driver.execute_script(
        "Array.from(document.querySelectorAll('a'))"
        ".find(el => el.textContent.trim() === '在线追踪签证申请状态').click();"
    )

    # 切换到 TRACKING 窗口
    WebDriverWait(driver, 10).until(lambda d: len(d.window_handles) > 1)
    found = False
    for _ in range(MAX_ATTEMPTS):
        for h in driver.window_handles:
            driver.switch_to.window(h)
            if "TRACKING" in driver.title.upper():
                found = True
                break
        if found:
            break
        time.sleep(3)
    if not found:
        logging.error("未找到包含 'TRACKING' 的窗口")
        raise RuntimeError("Window switch failed")

    # 等待主表单加载
    WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.ID, "AppRefNo")))
    driver.find_element(By.ID, "AppRefNo").send_keys(APP_REF)
    driver.find_element(By.NAME, "LastName").send_keys(LAST_NAME)

    # 验证码循环
    for attempt in range(1, MAX_ATTEMPTS + 1):
        time.sleep(2)
        png = driver.find_element(By.ID, "CaptchaImage").screenshot_as_png
        code = ocr.classification(png).strip().upper()
        logging.info(f"尝试第{attempt}次识别验证码：{code}")
        inp = driver.find_element(By.ID, "CaptchaInputText")
        inp.clear()
        inp.send_keys(code)
        driver.find_element(By.ID, "submitButton").click()
        time.sleep(2)
        if not driver.find_elements(By.CSS_SELECTOR, ".validation-summary-errors li"):
            logging.info("验证码通过")
            break
        logging.warning("验证码错误，刷新重新获取")
    else:
        raise RuntimeError("多次尝试后验证码仍然错误")

    # 提取状态
    found = False
    status_msg = ""
    for attempt in range(1, MAX_ATTEMPTS + 1):
        for b in driver.find_elements(By.TAG_NAME, "b"):
            if APP_REF in b.text:
                status_msg = b.text
                logging.info(f"Status: {status_msg}")
                found = True
                break
        if found:
            break
        logging.warning(f"Attempt {attempt}: status info not found, retrying…")
        time.sleep(3)
    if not found:
        logging.error(f"Status info not found after {MAX_ATTEMPTS} attempts")
        raise RuntimeError("Status info not found")

    # 发送邮件通知
    if SEND_EMAIL and "under process at Embassy/Consulate of Hungary" not in status_msg:
        import smtplib
        from email.mime.text import MIMEText

        # 邮件内容
        timestamp = time.strftime("%Y-%m-%d %H:%M:%S")
        timestamp = time.strftime("%Y-%m-%d %H:%M:%S")
        message = f"""
        <html>
          <body>
            <p><strong>Application Reference Number:</strong> {APP_REF}</p>
            <p><strong>Last Name:</strong> {LAST_NAME}</p>
            <p><strong>Status:</strong> {status_msg}</p>
            <p><strong>Checked at:</strong> {timestamp}</p>
          </body>
        </html>
        """
        msg = MIMEText(message, "html", "utf-8")
        msg["Subject"] = "Visa Application Status Update"
        msg["From"] = f"{SENDER} <{SENDER_EMAIL}>"
        msg["To"] = RECEIVER_EMAIL

        try:
            server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT, timeout=30)
            server.starttls()
            server.login(SENDER_EMAIL, SENDER_PASSWORD)
            server.sendmail(SENDER_EMAIL, RECEIVER_EMAIL, msg.as_string())
            logging.info("邮件已发送")
            server.quit()
        except smtplib.SMTPException as e:
            logging.error(f"邮件发送失败: {e}")
            raise
        except Exception as e:
            logging.exception("发送邮件时发生未知错误")
            raise

except Exception as e:
    logging.exception("脚本执行出错")
    traceback.print_exc()
finally:
    if driver:
        driver.quit()
