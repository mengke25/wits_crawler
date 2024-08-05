import os
import json
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException, StaleElementReferenceException, UnexpectedAlertPresentException, TimeoutException
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
import glob
import zipfile

def create_empty_zip(filepath):
    with zipfile.ZipFile(filepath, 'w') as zipf:
        pass


# 获取当前脚本所在的目录
current_dir = os.path.dirname(__file__)

# 构建config.json的相对路径
config_path = os.path.join(current_dir, 'config.json')

# 读取配置文件
with open(config_path, 'r') as f:
    config = json.load(f)

# 获取公共配置
root_path = config['common']['root_path']

# 获取当前脚本文件名并提取索引
script_name = os.path.basename(__file__)
subunit_index = int(script_name.replace("subunit", "").replace(".py", ""))
download_path = os.path.join(root_path, 'subunit', f'u{subunit_index}')

# 获取当前subunit的账号密码配置
account_config = next(item for item in config['accounts'] if item["index"] == subunit_index)
send_keys_account = account_config['send_keys_account']
send_keys_passwords = account_config['send_keys_passwords']

# 获取当前subunit的国家配置
subunit_config = next(item for item in config['subunits'] if item["index"] == subunit_index)
exporters = subunit_config['exporters']
importers = subunit_config['importers']

def main():
    def process_export_import(exporter, importer):
        while True:
            try:
                # 设置新的下载路径
                new_filename = f"{exporter}_{importer}.zip"
                new_filepath = os.path.join(download_path, new_filename)

                # 检查文件是否已经存在
                if os.path.exists(new_filepath):
                    print(f"文件 {new_filepath} 已存在，跳过此国家对")
                    return  # 文件存在，跳过此组合

                # Edge options
                options = webdriver.EdgeOptions()
                options.use_chromium = True

                # 修改默认下载路径
                prefs = {
                    "download.default_directory": download_path,
                    "download.prompt_for_download": False,
                    "download.directory_upgrade": True,
                    "safebrowsing.enabled": True
                }
                options.add_experimental_option("prefs", prefs)

                # 初始化WebDriver
                driver = webdriver.Edge(options=options)
                url = 'https://wits.worldbank.org/WITS/WITS/Restricted/Login.aspx'
                driver.get(url)

                # 等待页面加载并输入登录信息
                wait = WebDriverWait(driver, 20)
                input1 = wait.until(EC.presence_of_element_located((By.XPATH, '/html/body/form/div[3]/div/div/div[2]/div[1]/div[1]/div/input')))
                input1.send_keys(send_keys_account)

                input2 = driver.find_element(By.XPATH, '/html/body/form/div[3]/div/div/div[2]/div[1]/div[2]/div/input')
                input2.send_keys(send_keys_passwords)

                loginbutton = driver.find_element(By.XPATH, '/html/body/form/div[3]/div/div/div[2]/div[1]/div[3]/input')
                loginbutton.click()

                print("账号已登录")

                # 等待高级查询(advanced query)页面加载并悬停
                hover_element = wait.until(EC.presence_of_element_located((By.XPATH, '/html/body/form/table/tbody/tr[2]/td[2]/div/div/div/nav/ul/li[2]')))
                actions = ActionChains(driver)
                actions.move_to_element(hover_element).perform()

                # 等待并点击下拉选项(UN COMTRADE DATA)
                trade_data_option = wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/form/table/tbody/tr[2]/td[2]/div/div/div/nav/ul/li[2]/ul/li[1]/a')))
                trade_data_option.click()
                time.sleep(3)

                # existing query
                input_button_xpath = '/html/body/form/table/tbody/tr[7]/td[2]/table/tbody/tr[4]/td/table/tbody/tr[2]/td[2]/div/select'
                input_button = wait.until(EC.element_to_be_clickable((By.XPATH, input_button_xpath)))
                input_button.click()
                time.sleep(2)
                input_button.send_keys(Keys.ARROW_DOWN)
                time.sleep(2)
                input_button.send_keys(Keys.ENTER)

                # proceed
                driver.find_element(By.XPATH, '/html/body/form/table/tbody/tr[7]/td[2]/table/tbody/tr[4]/td/table/tbody/tr[5]/td[2]/input[1]').click()
                time.sleep(8)

                print("数据需求表单已登入")

                # 选择reporter
                reporter_button = driver.find_element(By.XPATH, '/html/body/form/table/tbody/tr[7]/td[2]/table/tbody/tr[5]/td[3]/div/a')
                reporter_button.click()
                time.sleep(5)

                try:
                    alert = wait.until(EC.alert_is_present())
                    alert.accept()
                except NoSuchElementException:
                    print("并无Alert弹出")
                time.sleep(5)

                iframe = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.TAG_NAME, 'iframe'))
                )
                driver.switch_to.frame(iframe)
                time.sleep(7)

                clearall = driver.find_element(By.XPATH, '/html/body/form/div[3]/table/tbody/tr[4]/td/table/tbody/tr[1]/td[2]/table/tbody/tr/td/div[1]/ul[1]/div/li[1]/span/a')
                clearall.click()
                time.sleep(2)

                plusimg = driver.find_element(By.XPATH, '/html/body/form/div[3]/table/tbody/tr[4]/td/table/tbody/tr[1]/td[1]/table/tbody/tr[1]/td[2]/a/img')
                plusimg.click()
                time.sleep(2)

                imputexporter = driver.find_element(By.XPATH, '/html/body/form/div[3]/table/tbody/tr[4]/td/table/tbody/tr[1]/td[1]/table/tbody/tr[3]/td[1]/textarea')
                imputexporter.click()
                imputexporter.send_keys(exporter)
                time.sleep(2)

                rightarrow = driver.find_element(By.XPATH, '/html/body/form/div[3]/table/tbody/tr[4]/td/table/tbody/tr[1]/td[1]/table/tbody/tr[3]/td[2]/input')
                rightarrow.click()
                time.sleep(2)

                proceedexporter = driver.find_element(By.XPATH, '/html/body/form/div[3]/table/tbody/tr[4]/td/table/tbody/tr[2]/td[2]/input[1]')
                proceedexporter.click()
                time.sleep(5)

                try:
                    alert = wait.until(EC.alert_is_present())
                    if alert.text == "Please select Country/Country Group":
                        alert.accept()
                        create_empty_zip(new_filepath)
                        print(f"选择国家错误，为 {exporter}_{importer} 创建空压缩包")
                        time.sleep(5)
                        driver.quit()
                        return True  # 跳过当前国家对
                except TimeoutException:
                    print("未出现'Please select Country/ Country Group'警告")

                print(f"出口国为: {exporter}")

                # 操作完成后，切回主文档
                driver.switch_to.default_content()

                # 选择partner
                partnerr_button = driver.find_element(By.XPATH, '/html/body/form/table/tbody/tr[7]/td[2]/table/tbody/tr[9]/td[3]/div/a')
                partnerr_button.click()
                time.sleep(6)

                iframe = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.TAG_NAME, 'iframe'))
                )
                driver.switch_to.frame(iframe)
                time.sleep(7)

                clearall = driver.find_element(By.XPATH, '/html/body/form/div[3]/table/tbody/tr[4]/td/table/tbody/tr[1]/td[2]/table/tbody/tr/td/div[1]/ul[1]/div/li[1]/span/a')
                clearall.click()
                time.sleep(2)

                plusimg = driver.find_element(By.XPATH, '/html/body/form/div[3]/table/tbody/tr[4]/td/table/tbody/tr[1]/td[1]/table/tbody/tr[1]/td[2]/a/img')
                plusimg.click()
                time.sleep(2)

                imputimporter = driver.find_element(By.XPATH, '/html/body/form/div[3]/table/tbody/tr[4]/td/table/tbody/tr[1]/td[1]/table/tbody/tr[3]/td[1]/textarea')
                imputimporter.click()
                imputimporter.send_keys(importer)
                time.sleep(2)

                rightarrow = driver.find_element(By.XPATH, '/html/body/form/div[3]/table/tbody/tr[4]/td/table/tbody/tr[1]/td[1]/table/tbody/tr[3]/td[2]/input')
                rightarrow.click()
                time.sleep(2)

                proceedimporter = driver.find_element(By.XPATH, '/html/body/form/div[3]/table/tbody/tr[4]/td/table/tbody/tr[2]/td[2]/input[1]')
                proceedimporter.click()
                time.sleep(5)

                try:
                    alert = wait.until(EC.alert_is_present())
                    if alert.text == "Please select Country/Country Group":
                        alert.accept()
                        create_empty_zip(new_filepath)
                        print(f"选择国家错误，为 {exporter}_{importer} 创建空压缩包")
                        time.sleep(5)
                        driver.quit()
                        return True  # 跳过当前国家对
                except TimeoutException:
                    print("未出现'Please select Country/ Country Group'警告")

                print(f"进口国为: {importer}")
                print(f"本次下载国家对为: {exporter} - {importer}")

                # 操作完成后，切回主文档

                driver.switch_to.default_content()

                # submit
                submit = driver.find_element(By.XPATH, '/html/body/form/table/tbody/tr[7]/td[2]/table/tbody/tr[15]/td[1]/input[2]')
                submit.click()
                print('数据需求表单已提交')

                try:
                    alert = wait.until(EC.alert_is_present())
                    alert.accept()
                except NoSuchElementException:
                    print("并无alert弹出")
                time.sleep(10)

                # 设置等待时间
                wait = WebDriverWait(driver, 20)

                # 尽可能优化XPath
                xpath_click = '/html/body/form/table/tbody/tr[7]/td[2]/table/tbody/tr[5]/td/table/tbody/tr[td[2]/input'
                xpath_final_click = '/html/body/form/table/tbody/tr[7]/td[2]/table/tbody/tr[6]/td/div/div/table/tbody/tr[2]/td[5]/input'

                driver.find_element(By.XPATH, '/html/body/form/table/tbody/tr[7]/td[2]/table/tbody/tr[td[5]/td/table/tbody/tr[td[2]/input').click()
                time.sleep(35)
                print("已经等待35秒")

                while True:
                    # 点击第一个元素
                    driver.find_element(By.XPATH, xpath_click).click()
                    time.sleep(10)  # 等待10秒
                    print("已经等待10秒")
                    # 检查最终元素是否出现
                    try:
                        final_element = WebDriverWait(driver, 1).until(EC.element_to_be_clickable((By.XPATH, xpath_final_click)))
                        final_element.click()  # 如果元素出现，点击它
                        print("数据已导出，继续后续流程")
                        break  # 跳出循环，继续执行其他代码
                    except:
                        print("数据尚未完全生成，请稍后")
                        continue  # 最终元素未出现，继续循环

                try:
                    alert = wait.until(EC.alert_is_present())
                    if alert.text == "Data is not available for this query.":
                        alert.accept()
                        create_empty_zip(new_filepath)
                        print(f"本次导出数据无效，为 {exporter}_{importer} 创建空压缩包")
                        time.sleep(3)
                        driver.quit()
                        return
                except:
                    print("本次导出数据有效（Data is available for this query.）")
                    time.sleep(5)
                    iframe = WebDriverWait(driver, 10).until(
                        EC.presence_of_element_located((By.TAG_NAME, 'iframe'))
                    )
                    driver.switch_to.frame(iframe)
                    time.sleep(6)

                    # 选中三个quantity相关的变量，加入到表单中
                    driver.find_element(By.XPATH, '/html/body/form/div[3]/table/tbody/tr[9]/td[1]/select/option[11]').click()
                    driver.find_element(By.XPATH, '/html/body/form/div[3]/table/tbody/tr[9]/td[3]/input[2]').click()
                    time.sleep(2)
                    driver.find_element(By.XPATH, '/html/body/form/div[3]/table/tbody/tr[9]/td[1]/select/option[12]').click()
                    driver.find_element(By.XPATH, '/html/body/form/div[3]/table/tbody/tr[9]/td[3]/input[2]').click()
                    time.sleep(2)
                    driver.find_element(By.XPATH, '/html/body/form/div[3]/table/tbody/tr[9]/td[1]/select/option[13]').click()
                    driver.find_element(By.XPATH, '/html/body/form/div[3]/table/tbody/tr[9]/td[3]/input[2]').click()
                    time.sleep(2)

                    # 点击Download
                    driver.find_element(By.XPATH, '/html/body/form/div[3]/table/tbody/tr[10]/td[1]/table/tbody/tr[2]/td[5]/input[2]').click()
                    time.sleep(10)

                    try:
                        alert = wait.until(EC.alert_is_present())
                        alert.accept()
                    except NoSuchElementException:
                        print("点击下载后，并无Alert弹出")

                    driver.switch_to.default_content()

                    driver.find_element(By.XPATH, '/html/body/form/table/tbody/tr[7]/td[2]/table/tbody/tr[td[5]/td/table/tbody/tr[td[2]/input').click()
                    time.sleep(6)
                    driver.find_element(By.XPATH, '/html/body/form/table/tbody/tr[7]/td[2]/table/tbody/tr[td[7]/td/div/div/table/tbody/tr[2]/td[3]/a/img').click()
                    time.sleep(6)

                    # 获取最新下载的文件
                    list_of_files = glob.glob(download_path + '/*')  # * 表示任意文件
                    latest_file = max(list_of_files, key=os.path.getmtime)  # 根据修改时间找到最新的文件

                    # 重命名文件
                    os.rename(latest_file, new_filepath)
                    print(f"文件已重命名为: {new_filepath}")
                    time.sleep(4)

                    driver.quit()
                    return  # 成功处理后退出循环
            except Exception as e:
                print(f"处理 {exporter} - {importer} 时出错: {e}")
                try:
                    driver.quit()
                except:
                    pass
                time.sleep(5)
                print("重新启动程序")

    for exporter in exporters:
        for importer in importers:
            if exporter == importer:
                continue  # 如果exporter和importer相同，跳过这个国家对

            process_export_import(exporter, importer)

if __name__ == "__main__":
    while True:
        try:
            main()
        except Exception as e:
            print(f"主程序发生错误: {e}")
            print("将在5秒后重新启动...")
            time.sleep(5)
