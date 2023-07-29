import xlrd
from selenium import webdriver
from selenium.webdriver.common.by import By
import urllib.request
import ddddocr
import time
from enum import Enum
from io import StringIO
from contextlib import redirect_stdout
import os
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException
import random
from datetime import datetime, timedelta


def correctFee(fee):
    if int(fee) < 0:
        return 0
    else:
        return fee
    
def clickCheckboxes(driver, id):
    spans = driver.find_elements(By.XPATH, "//input[@id='%s']/following-sibling::span" % id)
    spans = spans[:-1]
    for span in spans:
        span.click()

class DiseaseType(Enum):
    THYROID_BENIGN = 1
    THYROID_MALIGNANT = 2
    THYROID_BENIGN_RETROSTERNAL = 3
    BREAST_MAGLINANT = 4
    BREAST_BENIGN = 5

def execute():
    xls_cnt = 0
    xls_file_name = None
    for file_name in os.listdir(os.path.join(os.getcwd(), os.pardir)):
        if file_name.endswith("xls"):
            xls_file_name = file_name
            xls_cnt += 1
    if xls_cnt > 1:
        print("确保当前目录下只有一个EXCEL表格")
        driver.close()
        exit()
    if not xls_file_name:
        print("没有在当前目录下找到EXCEL表格")
        driver.close()
        exit()

    print("开启加载Chrome")
    driver = webdriver.Chrome()
    print("加载Chrome完毕")
    driver.get("https://quality.ncis.cn/report-disease/drgs")
    close = driver.find_elements(By.XPATH, "//div[@id='thediv']/div/i")
    if close:
        close[0].click()

    enter=driver.find_element(By.XPATH, '//span[text()="国家单病种质量管理与控制平台"]')
    enter.click()

    graph_element = WebDriverWait(driver, 20).until(EC.visibility_of_element_located((By.CLASS_NAME, "code-img")))
    time.sleep(1)
    img_dir = os.path.join(os.getcwd(), "img")
    if not os.path.isdir(img_dir):
        os.makedirs(img_dir)
    graph_url = graph_element.get_attribute("src")
    graph_path = os.path.join(img_dir, "%s.gif" % datetime.now().strftime("%Y-%m-%d-%H-%M-%S"))
    urllib.request.urlretrieve(graph_url, graph_path)
    with open(graph_path, 'rb') as f:
        img_bytes = f.read()
    output_string = StringIO()
    with redirect_stdout(output_string):
        ocr = ddddocr.DdddOcr()
        security_code = ocr.classification(img_bytes)

    with open(os.path.join(os.getcwd(), os.pardir, 'password.txt'), 'r') as f:
        lines = f.readlines()
    username = lines[0].strip()
    password = lines[1].strip()

    username_input = driver.find_element(By.XPATH, '//input[@placeholder="账户"]')
    password_input = driver.find_element(By.XPATH, '//input[@type="password"]')
    security_input = driver.find_element(By.XPATH, '//input[@placeholder="验证码"]')
    username_input.send_keys(username)
    password_input.send_keys(password)
    security_input.send_keys(security_code)

    login = driver.find_element(By.XPATH, '//span[text()="登录"]/parent::button')
    login.click()
    try:
        print("如自动填充验证码失败，请自行输入验证码并登录，无需在本命令行窗口敲回车。")
        WebDriverWait(driver, 100).until(EC.invisibility_of_element(login))
    except Exception as e:
        driver.close()
        exit()

    driver.get("https://quality.ncis.cn/report-disease/drgs")
    time.sleep(1)
    no_hint = driver.find_element(By.XPATH, '//label[contains(text(), "不再提示")]/span/input')
    no_hint.click()
    close_hint = driver.find_element(By.CLASS_NAME, 'ivu-icon.ivu-icon-ios-close')
    close_hint.click()

    workbook = xlrd.open_workbook(os.path.join(os.getcwd(), os.pardir, xls_file_name))
    sheet = workbook.sheet_by_index(0)
    data = []
    for nrow in range(sheet.nrows):
        data.append(sheet.row_values(nrow))
    data.pop(0)

    index = 0
    for row in data:
        index += 1
        driver.get("https://quality.ncis.cn/report-disease/drgs")
        time.sleep(1)

        disease = row[12]
        print("----------------------------------------------------")
        print("%d. 正在处理患者 (%s，%s) 数据" % (index, row[6].replace(" ", ""), disease))

        if "胸骨后甲状腺良性" in disease:
            kind = DiseaseType.THYROID_BENIGN_RETROSTERNAL
            driver.find_element(By.XPATH, '//li[contains(text(),"其他疾病/手术")]').click()
            driver.find_elements(By.XPATH, '//div[contains(text(), "甲状腺结节（手术治疗）")]/parent::div/div')[1].click()
        elif "甲状腺良性" in disease or "甲状腺结节" in disease:
            kind = DiseaseType.THYROID_BENIGN
            driver.find_element(By.XPATH, '//li[contains(text(),"其他疾病/手术")]').click()
            driver.find_elements(By.XPATH, '//div[contains(text(), "甲状腺结节（手术治疗）")]/parent::div/div')[1].click()
        elif "甲状腺恶性" in disease:
            kind = DiseaseType.THYROID_MALIGNANT
            driver.find_element(By.XPATH, '//li[contains(text(),"肿瘤(手术治疗)")]').click()
            driver.find_elements(By.XPATH, '//div[contains(text(), "甲状腺癌（手术治疗）")]/parent::div/div')[1].click()
        elif "乳腺恶性" in disease:
            kind = DiseaseType.BREAST_MAGLINANT
            driver.find_element(By.XPATH, '//li[contains(text(),"肿瘤(手术治疗)")]').click()
            driver.find_elements(By.XPATH, '//div[contains(text(), "乳腺癌（手术治疗）")]/parent::div/div')[1].click()
        elif "乳腺良性" in disease or "乳腺腺病" in disease or "乳房良性" in disease:
            kind = DiseaseType.BREAST_BENIGN
            driver.find_element(By.XPATH, '//li[contains(text(),"其他疾病/手术")]').click()
            driver.find_elements(By.XPATH, '//div[contains(text(), "围手术期预防感染")]/parent::div/div')[1].click()
        else:
            print("*** 暂不支持自动录入的疾病类型：%s" % disease)
            continue

        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, "myiframe")))
        driver.switch_to.frame("myiframe")
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, "submit")))
        
        fillSuccess = True
        try:
            # 质控医师
            quality_ctrl_doctor_input = driver.find_element(By.ID, "create_CM_1")
            quality_ctrl_doctor_input.send_keys(row[0].replace(" ", ""))

            # 质控护士
            quality_ctrl_nurse_input = driver.find_element(By.ID, "create_CM_2")
            quality_ctrl_nurse_input.send_keys(row[1].replace(" ", ""))

            # 主治医师
            attending_doctor_input = driver.find_element(By.ID, "create_CM_3")
            attending_doctor_input.send_keys(row[2].replace(" ", ""))

            # 责任护士
            primary_nurse_input = driver.find_element(By.ID, "create_CM_4")
            primary_nurse_input.send_keys(row[3].replace(" ", ""))

            # 上报科室
            submit_department_input = driver.find_element(By.ID, "create_CM_186")
            submit_department_input.send_keys(row[5])

            # 患者病案号
            patient_case_number_input = driver.find_element(By.ID, "create_CM_5")
            patient_case_number_input.send_keys(row[7])

            # 患者身份证号
            patient_id_number_input = driver.find_element(By.ID, "create_CM_6")
            patient_id_number_input.send_keys(row[8])

            # 主要诊断ICD-10四位亚目编码与名称
            try:
                primary_diagnosis_index_4_select = driver.find_element(By.ID, "create_CM_7")
                primary_diagnosis_index_4_select.send_keys(row[11].split(".")[0])
            except NoSuchElementException:
                pass

            # 主要诊断ICD-10六位临床扩展编码与名称
            try:
                primary_diagnosis_index_6_select = driver.find_element(By.ID, "create_CM_8")
                primary_diagnosis_index_6_select.send_keys(row[11])
            except NoSuchElementException:
                pass

            # 主要手术操作栏中提取ICD-9-CM-3四位亚目编码与名称
            primary_surgery_index_4_select = driver.find_element(By.ID, "create_CM_9")
            match kind:
                case DiseaseType.THYROID_BENIGN:
                    primary_surgery_index_4 = "06.2"
                case DiseaseType.THYROID_MALIGNANT:
                    primary_surgery_index_4 = "06.4"
                case DiseaseType.THYROID_BENIGN_RETROSTERNAL:
                    primary_surgery_index_4 = "06.5"
                case DiseaseType.BREAST_MAGLINANT:
                    primary_surgery_index_4 = "85.20"
                case DiseaseType.BREAST_BENIGN:
                    primary_surgery_index_4 = "乳房组织切除术"
                case _:
                    print("*** 未知疾病种类")
            primary_surgery_index_4_select.send_keys(primary_surgery_index_4)

            # 主要手术操作栏中提取ICD-9-CM-3六位临床扩展编码与名称
            primary_surgery_index_6_select = driver.find_element(By.ID, "create_CM_10")
            match kind:
                case DiseaseType.THYROID_BENIGN:
                    primary_surgery_index_6 = "06.2x00"
                case DiseaseType.THYROID_MALIGNANT:
                    primary_surgery_index_6 = "06.4x00"
                case DiseaseType.THYROID_BENIGN_RETROSTERNAL:
                    primary_surgery_index_6 = "06.5000"
                case DiseaseType.BREAST_MAGLINANT:
                    primary_surgery_index_6 = "85.2000"
                case DiseaseType.BREAST_BENIGN:
                    primary_surgery_index_6 = "85.2100x004"
                case _:
                    print("*** 未知疾病种类")
            primary_surgery_index_6_select.send_keys(primary_surgery_index_6)

            # 是否出院后31天内重复住院
            repeat_hospitalization_input = driver.find_elements(By.XPATH, '//input[@id="create_CM_11"]/parent::div/span')[2]
            repeat_hospitalization_input.click()

            # 患者性别
            if "女" in row[10]:
                driver.find_element(By.XPATH, '//span[@key="F"]').click()
                isMale = False
            else:
                driver.find_element(By.XPATH, '//span[@key="M"]').click()
                isMale = True

            # 患者体重
            if isMale:
                weight = 65 + random.randint(-5, 15)
            else:
                weight = 50 + random.randint(-5, 10)
            patient_height_input = driver.find_element(By.ID, "create_CM_15")
            patient_height_input.send_keys(weight)

            # 患者身高
            if isMale:
                if weight >= 75:
                    height = 170 + random.randint(0, 15)
                else:
                    height = 170 + random.randint(-5, 5)
            else:
                if weight >= 55:
                    height = 150 + random.randint(10, 15)
                else:
                    height = 150 + random.randint(0, 10)
            patient_weight_input = driver.find_element(By.ID, "create_CM_227")
            patient_weight_input.send_keys(height)

            # 入院日期时间
            js = """
                var date = document.getElementById(arguments[0]);
                date.readOnly = false;
                """
            driver.execute_script(js, "create_CM_16")
            admission_date_input = driver.find_element(By.ID, "create_CM_16")
            if (isinstance(row[13], str)):
                admission_date_obj = datetime.strptime(row[13], "%Y%m%d%H:%M:%S")
            else:
                admission_date_obj = xlrd.xldate_as_datetime(row[13], 0)
                delta = timedelta(hours=random.randint(8,9), minutes=random.randint(1,29))
                admission_date_obj += delta
            admission_date_input.send_keys(admission_date_obj.strftime('%Y-%m-%d %H:%M'))

            # 出院日期时间
            driver.execute_script(js, "create_CM_17")
            discharge_date_input = driver.find_element(By.ID, "create_CM_17")
            if (isinstance(row[14], str)):
                discharge_date_obj = datetime.strptime(row[14], "%Y%m%d%H:%M:%S")
            else:
                discharge_date_obj = xlrd.xldate_as_datetime(row[14], 0)
                delta = timedelta(hours=random.randint(15,17), minutes=random.randint(1,59))
                discharge_date_obj += delta
            discharge_date_input.send_keys(discharge_date_obj.strftime('%Y-%m-%d %H:%M'))

            # 手术开始时间
            driver.execute_script(js, "create_CM_24")
            surgery_begin_time_input = driver.find_element(By.ID, "create_CM_24")
            surgery_begin_time_obj = admission_date_obj.replace(hour = 0, minute = 0, second = 0)
            if (row[13] == row[14]): #日间手术
                delta_day = 0
                delta_hour = random.randint(10, 13)
                delta_minute = random.randint(0, 59)
            else:
                delta_day = random.randint(1,2)
                delta_hour = random.randint(9, 17)
                delta_minute = random.randint(0, 59)
            delta = timedelta(days = delta_day, hours = delta_hour, minutes = delta_minute)
            surgery_begin_time_obj += delta
            surgery_begin_time_input.send_keys(surgery_begin_time_obj.strftime('%Y-%m-%d %H:%M'))

            # 手术结束时间
            driver.execute_script(js, "create_CM_25")
            surgery_end_time_input = driver.find_element(By.ID, "create_CM_25")
            delta_hour = 1
            delta_minute = random.randint(15, 50)
            if kind == DiseaseType.BREAST_MAGLINANT:
                delta_hour = 2
            elif kind == DiseaseType.BREAST_BENIGN:
                delta_hour = 0
                delta_minute = random.randint(30, 45)
            delta = timedelta(hours = delta_hour, minutes = delta_minute)
            surgery_end_time_obj = surgery_begin_time_obj + delta
            surgery_end_time_input.send_keys(surgery_end_time_obj.strftime('%Y-%m-%d %H:%M'))

            # 费用支付方式
            payment_method_select = driver.find_element(By.ID, "create_CM_28")
            payment_method_select.send_keys(row[15])

            # 收入住院途径
            admission_route_select = driver.find_element(By.ID, "create_CM_29")
            admission_route_select.send_keys(row[16])

            # 到院交通工具
            transportation_method_select = driver.find_element(By.ID, "create_CM_30")
            transportation_method_select.send_keys(random.choice(["私家车", "出租车", "其它"]))

            # 离院方式选择
            leave_hospital_method_input = driver.find_element(By.ID, "create_CM_79")
            leave_hospital_method_input.send_keys("医嘱离院")

            # 住院总费用
            total_cost_input = driver.find_element(By.ID, "create_CM_98")
            total_cost_input.send_keys(correctFee(row[19]))

            # 住院总费用中自付金额
            self_cost_input = driver.find_element(By.ID, "create_CM_99")
            self_fee = row[20]
            if float(row[20]) > float(row[19]):
                self_fee = row[19]
            self_cost_input.send_keys(correctFee(self_fee))

            # 一般医疗服务费
            general_medical_service_fee_input = driver.find_element(By.ID, "create_CM_100")
            general_medical_service_fee_input.send_keys(correctFee(row[21]))

            # 一般治疗操作费
            general_treatment_operation_fee_input = driver.find_element(By.ID, "create_CM_101")
            general_treatment_operation_fee_input.send_keys(correctFee(row[22]))

            # 护理费
            nursing_fee_input = driver.find_element(By.ID, "create_CM_102")
            nursing_fee_input.send_keys(correctFee(row[23]))

            # 综合医疗服务类其他费用
            comprehensive_medical_service_other_fee_input = driver.find_element(By.ID, "create_CM_103")
            comprehensive_medical_service_other_fee_input.send_keys("0")

            # 病理诊断费
            pathology_diagnosis_fee_input = driver.find_element(By.ID, "create_CM_104")
            pathology_diagnosis_fee_input.send_keys(correctFee(row[24]))

            # 实验室诊断费
            laboratory_diagnosis_fee_input = driver.find_element(By.ID, "create_CM_105")
            laboratory_diagnosis_fee_input.send_keys(correctFee(row[25]))

            # 影像学诊断费
            diagnosis_imaging_fee_input = driver.find_element(By.ID, "create_CM_106")
            diagnosis_imaging_fee_input.send_keys(correctFee(row[26]))

            # 临床诊断项目费
            clinical_diagnosis_program_fee_input = driver.find_element(By.ID, "create_CM_107")
            clinical_diagnosis_program_fee_input.send_keys(correctFee(row[27]))

            # 非手术治疗项目费
            non_surgical_treatment_program_fee_input = driver.find_element(By.ID, "create_CM_108")
            non_surgical_treatment_program_fee_input.send_keys(correctFee(row[28]))

            # 临床物理治疗费
            clinical_physiotherapy_fee_input = driver.find_element(By.ID, "create_CM_109")
            clinical_physiotherapy_fee_input.send_keys(correctFee(row[29]))

            # 手术治疗费
            surgical_treatment_fee_input = driver.find_element(By.ID, "create_CM_110")
            surgical_treatment_fee_input.send_keys(correctFee(row[30]))

            # 麻醉费
            anesthesia_fee_input = driver.find_element(By.ID, "create_CM_111")
            anesthesia_fee_input.send_keys(correctFee(row[31]))

            # 手术费
            surgery_fee_input = driver.find_element(By.ID, "create_CM_112")
            surgery_fee_input.send_keys(correctFee(row[32]))

            # 康复费
            rehabilitation_fee_input = driver.find_element(By.ID, "create_CM_113")
            rehabilitation_fee_input.send_keys(correctFee(row[33]))

            # 中医治疗费
            tcm_fee_input = driver.find_element(By.ID, "create_CM_114")
            tcm_fee_input.send_keys(correctFee(row[34]))

            # 西药费
            western_medicine_fee_input = driver.find_element(By.ID, "create_CM_115")
            western_medicine_fee_input.send_keys(correctFee(row[35]))

            # 抗菌药物费
            antibacterial_drug_fee_input = driver.find_element(By.ID, "create_CM_116")
            antibacterial_drug_fee_input.send_keys(correctFee(row[36]))

            # 中成药费
            proprietary_chinese_medicine_fee_input = driver.find_element(By.ID, "create_CM_117")
            proprietary_chinese_medicine_fee_input.send_keys(correctFee(row[37]))

            # 中草药费
            chinese_herbal_medicine_fee_input = driver.find_element(By.ID, "create_CM_118")
            chinese_herbal_medicine_fee_input.send_keys(correctFee(row[38]))

            # 血费
            blood_cost_input = driver.find_element(By.ID, "create_CM_119")
            blood_cost_input.send_keys(correctFee(row[39]))

            # 白蛋白类制品费
            albumin_product_fee_input = driver.find_element(By.ID, "create_CM_120")
            albumin_product_fee_input.send_keys(correctFee(row[40]))

            # 球蛋白类制品费
            globulin_product_input = driver.find_element(By.ID, "create_CM_121")
            globulin_product_input.send_keys(correctFee(row[41]))

            # 凝血因子类制品费
            coagulation_factor_product_fee_input = driver.find_element(By.ID, "create_CM_122")
            coagulation_factor_product_fee_input.send_keys(correctFee(row[42]))

            # 细胞因子类制品费
            cytokine_product_fee_input = driver.find_element(By.ID, "create_CM_123")
            cytokine_product_fee_input.send_keys(correctFee(row[43]))

            # 检查用一次性医用材料费
            disposable_medical_metrials_for_examination_fee_input = driver.find_element(By.ID, "create_CM_124")
            disposable_medical_metrials_for_examination_fee_input.send_keys(correctFee(row[44]))

            # 治疗用一次性医用材料费
            disposable_medical_metrials_for_treatment_fee_input = driver.find_element(By.ID, "create_CM_125")
            disposable_medical_metrials_for_treatment_fee_input.send_keys(correctFee(row[45]))

            # 手术用一次性医用材料费
            disposable_medical_metrials_for_surgery_fee_input = driver.find_element(By.ID, "create_CM_126")
            disposable_medical_metrials_for_surgery_fee_input.send_keys(correctFee(row[46]))

            # 其他费
            other_fee_input = driver.find_element(By.ID, "create_CM_127")
            other_fee_input.send_keys(correctFee(row[47]))

            # 手术野皮肤准备常用方法选择
            driver.find_element(By.ID, "create_CM_72").send_keys("剪刀清除毛发")

            # 使用含抗菌剂缝线
            driver.find_element(By.ID, "create_CM_73").send_keys("不使用")

            # 手术切口类别的选择
            driver.find_element(By.ID, "create_CM_74").send_keys("无法确定或无记录")

            # 手术切口愈合情况的选择
            driver.find_element(By.ID, "create_CM_75").send_keys("无法确定或无记录")

            # 是否使用预防性抗菌药物（否）
            driver.find_elements(By.XPATH, "//input[@id='create_CM_32']/parent::div/span")[2].click()

            # 是否有手术并发症（否）
            driver.find_elements(By.XPATH, "//input[@id='create_CM_52']/parent::div/span")[2].click()

            # 患者是否对服务的体验与评价
            driver.find_elements(By.XPATH, "//input[@id='create_CM_85']/parent::div/span")[2].click()

            # 手术野皮肤准备常用方法选择/清洁
            driver.find_element(By.XPATH, "//select[@id='create_CM_72']").send_keys("清洁")

            # 使用含抗菌剂（三氯生）缝线/不使用
            driver.find_element(By.XPATH, "//select[@id='create_CM_73']").send_keys("不使用")

            # 手术切口类别的选择
            driver.find_element(By.XPATH, "//select[@id='create_CM_74']").send_keys("Ⅰ类切口")

            # 手术切口愈合情况的选择
            driver.find_element(By.XPATH, "//select[@id='create_CM_75']").send_keys("甲级愈合")

            match kind:
                case DiseaseType.BREAST_BENIGN:
                    driver.find_element(By.ID, "create_PIP_230").send_keys("日间手术")
                case DiseaseType.BREAST_MAGLINANT:
                    # 是否为T1-2,N0M0乳腺癌
                    driver.find_elements(By.XPATH, "//input[@id='create_26']/parent::div/span")[2].click()

                    # 是否乳腺癌治疗前TNM临床分期
                    driver.find_elements(By.XPATH, "//input[@id='create_35']/parent::div/span")[2].click()

                    # 是否乳腺癌手术治疗
                    driver.find_elements(By.XPATH, "//input[@id='create_49']/parent::div/span")[1].click()
                    driver.find_elements(By.XPATH, "//div[@id='item51']/span")[5].click()

                    # 术后病理报告记录，是否有肿瘤大小、切缘、脉管浸润（否）
                    driver.find_elements(By.XPATH, "//input[@id='create_63']/parent::div/span")[2].click()

                    # 术后病理报告记录，是否侵犯皮肤和/或胸壁（否）
                    driver.find_elements(By.XPATH, "//input[@id='create_64']/parent::div/span")[2].click()

                    # 术后病理报告记录，是否有检查淋巴结组数（否）
                    driver.find_elements(By.XPATH, "//input[@id='create_65']/parent::div/span")[2].click()

                    # 术后病理报告记录，是否有免疫组化检测内容（否）
                    driver.find_elements(By.XPATH, "//input[@id='create_12']/parent::div/span")[2].click()

                    # 术后病理报告记录，有无病理类型分级（否）
                    driver.find_elements(By.XPATH, "//input[@id='create_68']/parent::div/span")[2].click()

                    # 病理分期要素：T(原发肿瘤),TX 原发肿瘤无法评估
                    driver.find_element(By.ID, "create_69").send_keys("TX 原发肿瘤无法评估")

                    # 病理分期要素：N(局部淋巴结),pN1
                    driver.find_element(By.ID, "create_70").send_keys("pN1")

                    # 病理分期要素：M(远处转移),M0 无临床或者影像学证据
                    driver.find_element(By.ID, "create_71").send_keys("M0 无临床或者影像学证据")

                    # 乳腺癌TNM病理分期结论：(IA期 T1 N0 M0）
                    driver.find_element(By.ID, "create_72").send_keys("IA期 T1 N0 M0")

                    # 术后患者规范放疗
                    driver.find_elements(By.XPATH, "//div[@id='item107']/span")[1].click()

                    # 雌激素受体ER的评价结果
                    driver.find_elements(By.XPATH, "//div[@id='item117']/span")[1].click()

                    # HER-2(受体蛋白)的评价结果
                    driver.find_elements(By.XPATH, "//div[@id='item120']/span")[1].click()

                    # 交与患者“出院小结”的副本告知患者出院时风险因素
                    clickCheckboxes(driver, 'create_CM_150')

                    # 出院健康教育与告知
                    clickCheckboxes(driver, 'create_CM_153')

                    # 出院时教育与随访
                    clickCheckboxes(driver, 'create_CM_154')

                case DiseaseType.THYROID_MALIGNANT:
                    # 甲状腺癌治疗前是否在进行TNM临床分期(否)
                    driver.find_elements(By.XPATH, "//input[@id='create_29']/parent::div/span")[2].click()

                    # 术前评估有无淋巴结转移（否）
                    driver.find_elements(By.XPATH, "//input[@id='create_50']/parent::div/span")[2].click()

                    # 是否甲状腺癌手术治疗（是）
                    driver.find_elements(By.XPATH, "//input[@id='create_17']/parent::div/span")[1].click()

                    # 甲状腺癌再次手术（否）
                    driver.find_elements(By.XPATH, "//input[@id='create_52']/parent::div/span")[3].click()

                    # 甲状腺癌手术治疗方式
                    driver.find_element(By.ID, "create_55").send_keys("无法确定,或者无记录")

                    # 是否为0-31天内非计划二次手术（否）
                    driver.find_elements(By.XPATH, "//input[@id='create_TC_1930']/parent::div/span")[2].click()

                    # 术后病理诊断(否)
                    driver.find_elements(By.XPATH, "//input[@id='create_106']/parent::div/span")[2].click()

                    # 是否输血(否)
                    driver.find_elements(By.XPATH, "//input[@id='create_175']/parent::div/span")[2].click()

                    # 是否进行甲状腺复发风险评估(是)
                    driver.find_elements(By.XPATH, "//input[@id='create_138']/parent::div/span")[2].click()

                    # 术前健康教育
                    clickCheckboxes(driver, 'create_CM_156')

                    # 术后健康教育
                    clickCheckboxes(driver, 'create_CM_157')

                    # 履行出院知情告知
                    clickCheckboxes(driver, 'create_CM_150')

                    # 出院带药
                    clickCheckboxes(driver, 'create_CM_151')

                    # 告知发生紧急意外情况或者疾病复发如何救治及前途径
                    clickCheckboxes(driver, 'create_CM_153')

                    # 出院时教育与随访
                    clickCheckboxes(driver, 'create_CM_154')
                
                case DiseaseType.THYROID_BENIGN | DiseaseType.THYROID_BENIGN_RETROSTERNAL:
                    # 临床表现（甲状腺内发现肿块）
                    driver.find_elements(By.XPATH, "//div[@id='item40']/span")[1].click()

                    # 治疗前是否进行甲状腺及淋巴结超声检查（否）
                    driver.find_elements(By.XPATH, "//div[@id='item43']/span")[2].click()

                    # 甲状腺结节再次手术（否）
                    driver.find_elements(By.XPATH, "//div[@id='item61']/span")[2].click()

                    # 甲状腺手术适应症的选择（临床考虑有恶变倾向或合并甲状腺癌高危因素）
                    driver.find_elements(By.XPATH, "//div[@id='item65']/span")[6].click()

                    # 手术治疗方式选择（甲状腺部分切除）
                    driver.find_element(By.ID, "create_68").send_keys("甲状腺部分切除")

                    # 术中甲状腺病灶最大直径的选择（无法确定或无记录）
                    driver.find_element(By.ID, "create_71").send_keys("无法确定,或者无记录")

                    # 是否进行术中快速活体组织病理学检查（否）
                    driver.find_elements(By.XPATH, "//div[@id='item75']/span")[2].click()

                    # 是否有手术后并发症（否）
                    driver.find_elements(By.XPATH, "//div[@id='itemCM_52']/span")[2].click()

                    # 是否进行术后病理学检查（否）
                    driver.find_elements(By.XPATH, "//div[@id='itemTN_2057']/span")[2].click()

                    # 术前健康教育
                    clickCheckboxes(driver, 'create_CM_156')

                    # 术后健康教育
                    clickCheckboxes(driver, 'create_CM_157')

                    # 履行出院知情告知
                    clickCheckboxes(driver, 'create_CM_150')

                    # 出院带药
                    clickCheckboxes(driver, 'create_CM_151')

                    # 告知发生紧急意外情况或者疾病复发如何救治及前途径
                    clickCheckboxes(driver, 'create_CM_153')

                    # 出院时教育与随访
                    clickCheckboxes(driver, 'create_CM_154')
                case _:
                    print("Invalid Disease Type")
        except Exception as e:
            fillSuccess = False
            input(" *** 自动填充失败，请自行填充数据并提交，并在本命令行窗口按回车键继续。\n *** 如需跳过该条数据，直接按回车。")

        if fillSuccess:
            submit = driver.find_element(By.ID, "submit")
            driver.execute_script("arguments[0].focus()", submit)
            input(" a.自动填充完毕，请检查相关内容。\n b.确认无误后，请自行提交，并在本命令行窗口按回车键继续。\n c.如需跳过该条数据（比如数据重复），直接按回车。")

        continue
                        
    driver.close()

if __name__=="__main__":
    execute()