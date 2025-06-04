from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
from selenium import webdriver
from selenium.webdriver.edge.service import Service as EdgeService
from selenium.webdriver.edge.options import Options as EdgeOptions
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.chrome.service import Service
from bs4 import BeautifulSoup
import pandas as pd
import time
from PIL import Image
import pytesseract
import re


edge_driver_path = r"C:\Program Files (x86)\Microsoft\EdgeCore\msedgedriver.exe"
headers = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/135.0.0.0 Safari/537.36 Edg/135.0.0.0",
    "Accept": "*/*",
    "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8,en-GB;q=0.7,en-US;q=0.6",
    "Referer": "https://dierketang.ccnu.edu.cn/dektz/login",
}

options = EdgeOptions()
options.add_argument("--inprivate")  # Edge 无痕模式
options.add_argument("--disable-notifications") # 禁用弹窗
options.add_argument("--disable-blink-features=AutomationControlled")  # 绕过自动化检测
service = EdgeService(executable_path=edge_driver_path)
driver = webdriver.Edge(service=service, options=options)
driver.maximize_window()
url = 'https://www.okcis.cn/search/'
driver.get(url)

def solve_captcha(image_path):
    # 打开并预处理图像
    img = Image.open(image_path).convert('L')
    img = img.point(lambda x: 0 if x < 200 else 255)

    text = pytesseract.image_to_string(img, config='--psm 6')
    return text


def extract_numbers(text):
    # 匹配模式：允许数字间存在任意非数字字符（如 "+", 误识别的符号等）
    pattern = r"""
        \d+                # 第一个数字
        [^\d]+             # 中间非数字字符（如加号、误识别符号）
        (\d+)              # 第二个数字（关键：用捕获组提取）
    """

    # 搜索匹配（跨符号匹配）
    match = re.search(pattern, text, re.VERBOSE)
    if not match:
        return 2, 1

    # 提取数字（兼容中间符号干扰）
    num1 = int(re.search(r'\d+', text).group())  # 提取第一个连续数字
    num2 = int(match.group(1))  # 提取第二个捕获组数字
    return num1, num2


def highlight_element(driver, element):
    """高亮显示元素（用于调试）"""
    driver.execute_script("arguments[0].style.border='3px solid red'", element)

# 文本输入
def input_text(driver, by, selector, text):
    try:
        # 确保定位到input元素
        element = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((by, selector))
        )
        print("元素信息验证通过:", element.get_attribute("outerHTML"))

        element.clear()
        element.send_keys(text)

    except Exception as e:
        print(f"输入失败: {e}")
        driver.save_screenshot("error.png")
        raise


# 登录
# 使用显式等待定位元素，确保其可见
element = WebDriverWait(driver, 10).until(
    EC.visibility_of_element_located((By.CSS_SELECTOR, "a[name='site-top-login-button']"))  # 替换为按钮的实际定位方式
)

# 创建ActionChains对象并执行悬停
actions = ActionChains(driver)
actions.move_to_element(element).perform()

# 可选：验证悬停后的效果，例如等待下拉菜单出现
try:
    WebDriverWait(driver, 15).until(
        EC.frame_to_be_available_and_switch_to_it(
            (By.CSS_SELECTOR, "iframe[name^='site-top-login-iframe']"))
    )
    print("成功切换至目标iframe")

    target_ul = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, "ul.inde_dlul_20141021"))
    )
    print("找到目标 <ul> 元素")

    form_element = target_ul.find_element(
        By.CSS_SELECTOR, "form[name='iframe-login-form']"
    )
    print("找到表单")
    username_input = form_element.find_element(By.ID, "uname")
    username_input.clear()
    username_input.send_keys("**************")

    password_input = form_element.find_element(By.ID, "pwd")
    password_input.clear()
    password_input.send_keys("***********")

    while True:
        img_element = WebDriverWait(form_element, 30).until(
            EC.visibility_of_element_located((By.ID, "setcode"))
        )
        time.sleep(1)
        img_element.click()
        time.sleep(1)
        img_element.screenshot("cpatcha.jpg")
        text = solve_captcha("cpatcha.jpg")
        num1, num2 = extract_numbers(text)
        print(num1, num2)
        num = num1 - num2
        print(num)
        password_input = form_element.find_element(By.ID, "yzm")
        password_input.clear()
        password_input.send_keys(f"{num} ")
        time.sleep(1)
        button = WebDriverWait(form_element, 20).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, ".index_dlli_sub"))
        )
        button.click()

        try:
            # 等待表单消失（最长等待5秒）
            WebDriverWait(target_ul, 2).until(
                EC.invisibility_of_element_located((By.CSS_SELECTOR, "form[name='iframe-login-form']"))
            )
            print("登录成功，退出循环")
            break
        except TimeoutException:
            print("登录失败，继续重试")

except TimeoutException:
    print("悬停后未检测到预期变化。")

driver.switch_to.default_content()
print("已切换回主文档")


# 限制条件，搜索
try:
    element = WebDriverWait(driver, 20).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, "#search-form-on input[type='text']"))
    )

    driver.execute_script("arguments[0].value = '';", element)  # 先通过JS清空
    element.send_keys("评估")
    print("输入成功，当前值:", element.get_attribute('value'))
except Exception as e:
    print(f"输入失败: {e}")
    driver.save_screenshot("error.png")
    raise

time.sleep(2)

form_element = driver.find_element(
        By.CSS_SELECTOR, "form[name='search-form-on']"
)
print("找到表单")
huazhong_button = WebDriverWait(driver, 15).until(
    EC.element_to_be_clickable(
        (By.XPATH, '//label[.//span[contains(text(), "华中区")] and .//input[@name="city-class-7"]]')
    )
)
huazhong_button.click()
time.sleep(2)
today_button = WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable(
        (By.XPATH, '//label[contains(@class, "sjs_qbf_20160118")]//span[text()="当天"]/..')
    )
)
today_button.click()
time.sleep(2)

button = WebDriverWait(driver, 20).until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, ".bh_g_anniu"))
)
button.click()

time.sleep(2)

i = 1

# 爬取并导出数据
try:
    # 关键步骤1：等待主框架加载
    WebDriverWait(driver, 15).until(
        EC.frame_to_be_available_and_switch_to_it(
            (By.CSS_SELECTOR, "iframe[name^='iframe0.']"))
    )
    print("成功切换至目标iframe")
    # 定位包含所有<li>的父级<ul>
    all_projects = []

    while True:
        try:
            print(i)
            print(i)
            print(i)
            if i == 21:
                WebDriverWait(driver, 15).until(
                    EC.frame_to_be_available_and_switch_to_it(
                        (By.CSS_SELECTOR, "iframe[name^='layui-layer-iframe10']"))
                )
                print("成功切换至目标iframe222222222222")

                target_ul = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.ID, "cons"))
                )
                print("找到目标 <ul> 元素222222222222222")

                form_element = target_ul.find_element(
                    By.CSS_SELECTOR, ".zhuce_ta_20140701"
                )
                print("找到表单222222222222222")

                # 定位验证码图片（等待元素加载）
                captcha_img = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located(
                        (By.XPATH, '//img[contains(@src, "checkUser_code_yanzheng_page.php")]'))
                )

                # 定位验证码输入框
                captcha_input = WebDriverWait(driver, 10).until(
                    EC.visibility_of_element_located((By.CLASS_NAME, "yanzheng_tex_20141021"))
                )

                # 输入验证码
                captcha_input.send_keys("1234")
                driver.switch_to.default_content()
                print("已切换回主文档22222222222222")


            time.sleep(5)
            project_list = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, '.tybl_list'))
            )
            print(project_list)

            time.sleep(2)

            # 在容器内获取所有有效ul元素
            projects = WebDriverWait(project_list, 10).until(
                EC.presence_of_all_elements_located((By.TAG_NAME, "ul"))
            )
            print(f"成功定位到{len(projects)}个项目")

            for project in projects:
                li_items = project.find_elements(By.TAG_NAME, "li")
                print(len(li_items))
                for item in li_items:
                    print(item.text)
                print(project)

                # 从<b>标签的rec属性解析数据
                b_tag = project.find_element(By.CSS_SELECTOR, "b.setwidth")
                rec_data = b_tag.get_attribute("rec").split("_")

                # 从<a>标签提取链接和标题
                a_tag = project.find_element(By.CSS_SELECTOR, "a[name='result-list-title']")
                detail_url = a_tag.get_attribute("href")
                title = a_tag.get_attribute("title")

                # 结构化输出
                project_data = {
                    "项目名称": title.split("招标计划")[0].strip(),
                    "公告类型": "招标计划" if "招标计划" in title else "其他",
                    "发布时间": b_tag.get_attribute("rec_jointime"),
                    "详情链接": detail_url,
                }
                # print(project_data)
                all_projects.append(project_data)

            print("22222222222222")
            df = pd.DataFrame(all_projects)
            df.to_excel("招标项目列表.xlsx", index=False, engine="openpyxl")
            print("数据导出成功！")

            # 尝试定位并点击下一页
            next_page = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable(
                    (By.XPATH, '//a[contains(@onclick, "loadnextpage") and text()="下一页"]'))
            )
            driver.execute_script("arguments[0].click();", next_page)
            i = i + 1
        except TimeoutException:
            print("已到最后一页")
            break




except Exception as e:
    print(f"全局异常: {str(e)}")
    driver.save_screenshot("error.png")
    raise
finally:
    driver.switch_to.default_content()






