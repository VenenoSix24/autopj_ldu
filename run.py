import time
import os
import subprocess # <--- 导入 subprocess 模块
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, TimeoutException, ElementClickInterceptedException

# --- 配置区 ---
# Chrome 可执行文件的完整路径 (根据你的系统修改)
CHROME_EXECUTABLE_PATH = r"C:\Program Files\Google\Chrome\Application\chrome.exe"
# 或者: CHROME_EXECUTABLE_PATH = r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"

# ChromeDriver 的完整路径 (推荐方式，假设 chromedriver 与脚本在同一目录)
script_dir = os.path.dirname(os.path.abspath(__file__))
CHROMEDRIVER_PATH = os.path.join(script_dir, 'chromedriver.exe')
# 或者，如果不在同一目录，请直接指定完整路径:
# CHROMEDRIVER_PATH = r'C:\path\to\your\webdriver\chromedriver.exe'

# 用于 Selenium 连接的调试端口 (保持不变)
DEBUGGING_PORT = 9222
DEBUGGER_ADDRESS = f'127.0.0.1:{DEBUGGING_PORT}'

# 独立用户数据目录 (路径可以自定义，确保父目录存在)
USER_DATA_DIR = r"C:\Selenium\ChromeProfile_Auto" # 使用新目录与手动启动区分

# 教务系统登录页面的 URL (！！非常重要：请替换为实际的登录页面 URL！！)
# 注意：这个 URL 应该是不带 vpn 前缀的原始登录地址，或者你确定能直接访问的 VPN 地址
LOGIN_PAGE_URL = "https://vpn.ldu.edu.cn/https/77726476706e69737468656265737421e8e44b8b693c6c45300d8db9d6562d/index" 
# 如果你知道登录后的课程列表页 URL，也可以保留它用于后续检查
COURSE_LIST_URL = 'https://vpn.ldu.edu.cn/https/77726476706e69737468656265737421e8e44b8b693c6c45300d8db9d6562d/student/teachingEvaluation/newEvaluation/index'

# 评语内容
COMMENT_TEXT = "无"
# 每个评价选项后和点击保存前的等待时间 (秒) - 至少需要5秒
WAIT_BEFORE_SAVE = 6
# Selenium 操作的默认等待超时时间 (秒)
DEFAULT_WAIT_TIMEOUT = 20
# 启动 Chrome 后等待的时间 (秒)
WAIT_AFTER_CHROME_LAUNCH = 5
# --- 配置区结束 ---


def launch_chrome_with_debugging(chrome_path, port, user_dir, url):
    """使用指定的参数启动 Chrome 浏览器"""
    print(f"正在尝试启动 Chrome...")
    print(f"  路径: {chrome_path}")
    print(f"  端口: {port}")
    print(f"  用户目录: {user_dir}")
    print(f"  目标URL: {url}")

    # 确保用户目录存在
    os.makedirs(user_dir, exist_ok=True)

    # 构建命令列表
    command = [
        chrome_path,
        f"--remote-debugging-port={port}",
        f"--user-data-dir={user_dir}",
        url  # 将 URL 作为最后一个参数传递给 Chrome
    ]
    try:
        # 使用 Popen 启动 Chrome，使其在后台运行
        subprocess.Popen(command)
        print(f"Chrome 启动命令已发送。等待 {WAIT_AFTER_CHROME_LAUNCH} 秒让浏览器启动...")
        time.sleep(WAIT_AFTER_CHROME_LAUNCH)
        print("请检查是否已打开 Chrome 窗口并导航到登录页面。")
        return True
    except FileNotFoundError:
        print(f"错误：找不到 Chrome 可执行文件，请检查路径配置：'{chrome_path}'")
        return False
    except Exception as e:
        print(f"启动 Chrome 时发生错误: {e}")
        return False

def setup_driver():
    """设置并连接到已运行的 Chrome 浏览器"""
    print("正在设置 Chrome WebDriver 并连接到已启动的实例...")
    chrome_options = Options()
    chrome_options.add_experimental_option("debuggerAddress", DEBUGGER_ADDRESS)
    service = Service(CHROMEDRIVER_PATH)
    try:
        driver = webdriver.Chrome(service=service, options=chrome_options)
        print(f"成功连接到运行在 {DEBUGGER_ADDRESS} 的 Chrome 实例。")
        # 连接成功后，验证一下是否在登录后的某个页面 (例如课程列表页)
        time.sleep(2) # 等待页面稳定
        if COURSE_LIST_URL not in driver.current_url and LOGIN_PAGE_URL not in driver.current_url:
             print(f"警告：连接成功，但当前页面既不是登录页也不是课程列表页。")
             print(f"当前URL: {driver.current_url}")
             print(f"请确保你已在浏览器中成功登录。")
        elif LOGIN_PAGE_URL in driver.current_url:
             print(f"警告：仍然在登录页面，请确保已手动登录成功！")

        return driver
    except Exception as e:
        # 区分连接失败和其他 WebDriver 错误
        if "failed to connect" in str(e) or "cannot connect" in str(e):
             print(f"连接 WebDriver 失败: {e}")
             print("请确认：")
             print(f"1. Chrome 是否已成功启动并监听在端口 {DEBUGGING_PORT}?")
             print(f"2. ChromeDriver 路径 '{CHROMEDRIVER_PATH}' 是否正确且版本兼容？")
             print("3. 是否有其他程序占用了该端口？")
        else:
            print(f"设置 WebDriver 时发生其他错误: {e}")
        return None

# --- evaluate_single_course 函数保持不变 ---
def evaluate_single_course(driver):
    """在单个课程评估页面执行评估操作"""
    print("  进入单个课程评估页面...")
    try:
        # 等待页面加载完成 (可以等待某个特定元素出现，例如第一个 radio button 或 h4 标题)
        WebDriverWait(driver, DEFAULT_WAIT_TIMEOUT).until(
            EC.presence_of_element_located((By.XPATH, "//input[@type='radio'][@value='A_优']"))
        )
        print("  页面加载完成。")

        # 1. 选择所有问题的 "优" 选项
        print("  选择所有问题的 '优' 选项...")
        excellent_radios = driver.find_elements(By.XPATH, "//input[@type='radio'][@value='A_优']")
        if not excellent_radios:
            print("  警告：未找到 '优' 选项的单选按钮。")
            return False # 如果找不到选项，则无法继续

        for radio in excellent_radios:
            try:
                # 使用 JavaScript 点击 (更可靠)
                driver.execute_script("arguments[0].click();", radio)
                time.sleep(0.1) # 短暂暂停，避免过快点击
            except ElementClickInterceptedException:
                 print(f"  警告：直接点击单选按钮 {radio.get_attribute('name')} 被拦截，尝试JS点击。")
                 try:
                      driver.execute_script("arguments[0].click();", radio)
                      time.sleep(0.1)
                 except Exception as js_e:
                     print(f"  错误：JS点击单选按钮 {radio.get_attribute('name')} 也失败: {js_e}")
            except Exception as e:
                print(f"  错误：点击单选按钮 {radio.get_attribute('name')} 时出错: {e}")

        print(f"  已尝试选择 {len(excellent_radios)} 个 '优' 选项。")

        # 2. 填写评语
        print(f"  填写评语 '{COMMENT_TEXT}'...")
        try:
            # 定位最后一个 textarea
            comment_box = driver.find_element(By.XPATH, "(//textarea)[last()]")
            comment_box.clear() # 清空可能存在的默认内容
            comment_box.send_keys(COMMENT_TEXT)
            print("  评语填写完成。")
        except NoSuchElementException:
            print("  警告：未找到评语文本框 (textarea)。")
        except Exception as e:
            print(f"  错误：填写评语时出错: {e}")
            return False

        # 3. 等待指定时间
        print(f"  等待 {WAIT_BEFORE_SAVE} 秒...")
        time.sleep(WAIT_BEFORE_SAVE)

        # 4. 点击保存按钮
        print("  点击 '保存' 按钮...")
        try:
            # 等待保存按钮变为可点击状态
            save_button = WebDriverWait(driver, DEFAULT_WAIT_TIMEOUT + 5).until(
                EC.element_to_be_clickable((By.ID, "savebutton"))
            )
            # 使用 JavaScript 点击 (更可靠)
            driver.execute_script("arguments[0].click();", save_button)
            print("  '保存' 按钮已点击。")
            # 等待页面跳转或刷新
            print("  等待页面响应...")
            time.sleep(5) # 等待几秒让页面跳转回列表页或处理完成
            return True
        except TimeoutException:
            print("  错误：等待 '保存' 按钮变为可点击状态超时。")
            return False
        except ElementClickInterceptedException:
            print("  错误：点击 '保存' 按钮被拦截。")
            try:
                print("  尝试使用JS点击 '保存' 按钮...")
                save_button = driver.find_element(By.ID, "savebutton")
                driver.execute_script("arguments[0].click();", save_button)
                print("  JS点击 '保存' 按钮成功。")
                print("  等待页面响应...")
                time.sleep(5)
                return True
            except Exception as js_save_e:
                print(f"  错误：JS点击 '保存' 按钮也失败: {js_save_e}")
                return False
        except Exception as e:
            print(f"  错误：点击 '保存' 按钮时出错: {e}")
            return False

    except TimeoutException:
        print("  错误：等待单个课程评估页面元素加载超时。")
        return False
    except Exception as e:
        print(f"  错误：在评估单个课程时发生意外错误: {e}")
        return False


# --- main 函数修改 ---
def main():
    """主函数，执行自动化评估流程"""

    # 步骤 1: 自动启动 Chrome
    if not launch_chrome_with_debugging(CHROME_EXECUTABLE_PATH, DEBUGGING_PORT, USER_DATA_DIR, LOGIN_PAGE_URL):
        return # 启动失败，退出脚本

    # 步骤 2: 提示用户手动登录
    print("\n***********************************************************")
    print("请在刚刚打开的 Chrome 浏览器窗口中手动完成登录操作。")
    print("登录成功后，请确保已进入教务系统主界面或课程列表页面。")
    input("完成后，请回到这里按 Enter 键继续执行自动评估...")
    print("***********************************************************\n")

    # 步骤 3: 连接到已登录的 Chrome 实例
    driver = setup_driver()
    if not driver:
        print("无法连接到浏览器，脚本终止。")
        return # 连接失败，退出脚本

    # --- 后续评估逻辑与之前相同 ---
    evaluated_count = 0
    max_retries_on_tab_click = 3 # 增加标签页点击重试次数

    while True:
        print("\n--------------------")
        print("正在课程列表页面查找未评估的课程...")
        retry_count = 0 # 重置重试计数器

        # 1. 检查是否在课程列表页，如果不是则尝试导航 (必须在登录后！)
        # 使用 startswith 可能更灵活
        if not driver.current_url.startswith(COURSE_LIST_URL.split('?')[0]):
            print(f"当前不在预期的课程列表页 (URL: {driver.current_url})")
            print(f"尝试导航到: {COURSE_LIST_URL}")
            try:
                driver.get(COURSE_LIST_URL)
                # 等待页面加载，可以等待标签栏出现
                WebDriverWait(driver, DEFAULT_WAIT_TIMEOUT).until(
                    EC.presence_of_element_located((By.ID, "myTab"))
                )
                time.sleep(1) # 稍微等待JS渲染
            except TimeoutException:
                print("错误：导航到课程列表页后，等待标签栏加载超时。")
                # 可能登录失效或 URL 错误
                break
            except Exception as nav_e:
                print(f"错误：导航到课程列表页失败: {nav_e}")
                # 可能登录失效或 URL 错误
                break

        # 2. ******** 确保 "课堂教师" 标签页被选中 ********
        while retry_count < max_retries_on_tab_click:
            try:
                print(f"  检查 '课堂教师' 标签页状态 (尝试 {retry_count + 1}/{max_retries_on_tab_click})...")
                teacher_tab_link_xpath = "//ul[@id='myTab']/li/a[contains(normalize-space(), '课堂教师')]"
                teacher_tab_link = WebDriverWait(driver, DEFAULT_WAIT_TIMEOUT).until(
                    EC.presence_of_element_located((By.XPATH, teacher_tab_link_xpath))
                )
                parent_li = driver.find_element(By.XPATH, f"{teacher_tab_link_xpath}/parent::li")
                is_active = 'active' in parent_li.get_attribute('class')

                if not is_active:
                    print("  '课堂教师' 标签页不是活动的，正在点击...")
                    clickable_teacher_tab_link = WebDriverWait(driver, DEFAULT_WAIT_TIMEOUT).until(
                        EC.element_to_be_clickable((By.XPATH, teacher_tab_link_xpath))
                    )
                    driver.execute_script("arguments[0].click();", clickable_teacher_tab_link)
                    print("  '课堂教师' 标签页已点击。")
                    time.sleep(3) # 增加等待时间以确保表格加载
                    parent_li_after_click = driver.find_element(By.XPATH, f"{teacher_tab_link_xpath}/parent::li")
                    if 'active' in parent_li_after_click.get_attribute('class'):
                         print("  确认 '课堂教师' 标签页已激活。")
                         break # 成功激活
                    else:
                         print("  警告：点击后 '课堂教师' 标签页仍未激活，将重试...")
                         retry_count += 1
                         time.sleep(2) # 重试前稍等
                else:
                    print("  '课堂教师' 标签页已经是活动的。")
                    break # 已经是活动的

            except TimeoutException:
                print(f"  错误：查找或等待 '课堂教师' 标签页超时 (尝试 {retry_count + 1}/{max_retries_on_tab_click})。")
                retry_count += 1
                time.sleep(2) # 重试前稍等
                driver.refresh() # 尝试刷新页面
                time.sleep(3)
            except Exception as e:
                print(f"  错误：在处理 '课堂教师' 标签页时发生异常: {e} (尝试 {retry_count + 1}/{max_retries_on_tab_click})。")
                retry_count += 1
                time.sleep(2) # 重试前稍等
                driver.refresh() # 尝试刷新页面
                time.sleep(3)

        if retry_count >= max_retries_on_tab_click:
             print("错误：多次尝试后仍无法激活 '课堂教师' 标签页，脚本终止。")
             break

        # 3. ******** 继续执行原来的查找和评估逻辑 ********
        try:
            print("  正在查找课程表格...")
            WebDriverWait(driver, DEFAULT_WAIT_TIMEOUT).until(
                EC.visibility_of_element_located((By.ID, "codetbody"))
            )
            time.sleep(1)

            rows = driver.find_elements(By.XPATH, "//tbody[@id='codetbody']/tr")
            if not rows:
                print("在 '课堂教师' 标签页未找到任何课程行。检查是否真的没有未评估课程了。")
                break

            found_unevaluated = False
            for i, row in enumerate(rows):
                try:
                    status_element = row.find_element(By.XPATH, ".//td[last()]//span[contains(@class, 'label-light') or contains(@class, 'label-success')]")
                    is_evaluated = 'label-success' in status_element.get_attribute('class')

                    if not is_evaluated:
                        print(f"找到未评估课程 (行 {i+1})。")
                        found_unevaluated = True

                        try:
                            course_name = row.find_element(By.XPATH, ".//td[5]").text
                            teacher_name = row.find_element(By.XPATH, ".//td[4]").text
                            print(f"  课程: {course_name}, 教师: {teacher_name}")
                        except NoSuchElementException:
                            print("  无法提取课程/教师名称。")

                        evaluate_button_xpath = ".//td[2]//button[contains(., '评估')]"
                        evaluate_button = WebDriverWait(row, 5).until(
                             EC.element_to_be_clickable((By.XPATH, evaluate_button_xpath))
                        )

                        print("  点击 '评估' 按钮...")
                        driver.execute_script("arguments[0].click();", evaluate_button)
                        time.sleep(3)

                        success = evaluate_single_course(driver)

                        if success:
                            evaluated_count += 1
                            print(f"课程评估成功。已完成 {evaluated_count} 门课程。")
                            break # 跳出 for 循环，让 while 循环重新开始查找
                        else:
                            print("评估此课程失败。将尝试下一门课程（如果存在）。")
                            if not driver.current_url.startswith(COURSE_LIST_URL.split('?')[0]):
                                print("  评估失败后不在列表页，导航回去...")
                                driver.get(COURSE_LIST_URL)
                                time.sleep(3)
                            continue

                except NoSuchElementException:
                    print(f"  行 {i+1} 结构异常或未找到状态/按钮元素，跳过。")
                    continue
                except TimeoutException:
                     print(f"  行 {i+1} 等待评估按钮超时，跳过。")
                     continue
                except Exception as e:
                    print(f"  处理行 {i+1} 时出错: {e}")
                    continue

            if not found_unevaluated:
                print("当前列表页面未找到状态为 '否' 的未评估课程。脚本结束。")
                break

        except TimeoutException:
            print("错误：等待课程列表表格 ('codetbody') 加载超时或变为可见超时。")
            break
        except Exception as e:
            print(f"在查找/处理课程列表时发生严重错误: {e}")
            break

    print("\n--------------------")
    print(f"自动化评估脚本执行完毕。共成功评估 {evaluated_count} 门课程。")
    print("你可以手动关闭 Chrome 浏览器窗口了。")
    # driver.quit() # 如果希望脚本自动关闭浏览器，取消此行注释，但这会关闭你登录的实例

if __name__ == "__main__":
    main()