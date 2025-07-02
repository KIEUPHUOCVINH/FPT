from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import  TimeoutException
import pandas as pd
import time
import os

dict_input = {
    "Mã số thuế": ["0304244470", "0304244471", "0304308445", "", "", "", "", "", ""],
    "Mã tra cứu": ["r08e17y79g", "r46jvxmvxg", "rzmwy1yo4g", "B1HEIRR8N0WP", "PZH_FWQ4BN3", "VBHKSL682918", "NII30XVQWNC", "MHPLO8W6EMD", "MIJ634K9JAD"],
    "URL": [
        "https://tracuuhoadon.fpt.com.vn/search.html",  
        "https://tracuuhoadon.fpt.com.vn/search.html",
        "https://tracuuhoadon.fpt.com.vn/search.html",
        "https://www.meinvoice.vn/tra-cuu/",
        "https://www.meinvoice.vn/tra-cuu/",
        "https://www.meinvoice.vn/tra-cuu/",
        "https://van.ehoadon.vn/TCHD?MTC=",
        "https://van.ehoadon.vn/TCHD?MTC=",
        "https://van.ehoadon.vn/TCHD?MTC="
    ]
}
df = pd.DataFrame(dict_input)
df.to_excel("input.xlsx", index=False)

# Mở Chrome với tùy chọn tải về
def doi_file_tai_xong(folder_path, timeout=60):
    """
    Kiểm tra và chờ đợi file tải về hoàn thành
    """  
    start_time = time.time()
    while time.time() - start_time < timeout:
        files = os.listdir(folder_path)
        
        downloading = any(file.endswith('.crdownload') for file in files)        
        if not downloading and any(file.endswith('.xml') for file in files):
            return True
        time.sleep(2)  # Chờ 2 giây trước khi kiểm tra lại
    
    return False

def doi_ten_file_crdownload(folder_path, new_ext=".xml"):
    for f in os.listdir(folder_path):
        if f.endswith(".crdownload"):
            base = f[:-11]  # bỏ .crdownload
            new_name = base + new_ext
            os.rename(os.path.join(folder_path, f), os.path.join(folder_path, new_name))
            print(f" Đã đổi tên file thành: {new_name}")
            return
def open_chrome():
    options = Options()
    options.add_argument("--start-maximized")
    options.add_experimental_option("prefs", {
        "download.default_directory": r"D:\RPA\duanFPT",
        "download.directory_upgrade": True,
        "download.prompt_for_download": False,
        "disable-popup-blocking": "true",
         "safebrowsing.enabled": True
    })
    return webdriver.Chrome(service=Service(), options=options)
# Hàm tra cứu
def tra_cuu_hoa_don(driver, url, mst, ma_tra_cuu):
    try:
        driver.get(url)
        time.sleep(5)
        if "fpt" in url:
            driver.find_element(By.XPATH, '//input[@placeholder="MST bên bán"]').send_keys(mst)
            driver.find_element(By.XPATH, '//input[@placeholder="Mã tra cứu hóa đơn"]').send_keys(ma_tra_cuu)
            driver.find_element(By.XPATH, '//button[contains(text(), "Tra cứu")]').click()
            print(f" FPT: {mst} - {ma_tra_cuu}")

        elif "meinvoice.vn" in url:
            driver.find_element(By.XPATH, '//*[@id="txtCode"]').send_keys(ma_tra_cuu)
            driver.find_element(By.ID, "btnSearchInvoice").click()
            print(f" MISA: {ma_tra_cuu}")

        elif "van.ehoadon.vn" in url:
            try:
                # Gửi mã tra cứu
                code_input = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.ID, "txtInvoiceCode"))
                )
                code_input.clear()
                code_input.send_keys(ma_tra_cuu)

                # Click nút tra cứu bằng JavaScript để tránh bị iframe che mất
                search_button = driver.find_element(By.ID, "Button1")
                driver.execute_script("arguments[0].click();", search_button)
                print(f" VAN: {ma_tra_cuu}")
            except Exception as e:
                print(f"  Lỗi tra cứu (evanhoadon.vn): {e}")
        else:
            print(f" Trang không hỗ trợ: {url}")

        time.sleep(3)
    except Exception as e:
        print(f" Lỗi tra cứu: {e}")

# Kiểm tra kết quả tìm kiếm
def kiem_tra_ket_qua(driver, url):
    wait = WebDriverWait(driver, 7)
    try:
        if "fpt" in url:
            try:
                wait = WebDriverWait(driver, 5)
                wait.until(EC.visibility_of_element_located((
                By.XPATH, '//div[@view_id="search:status"]//span[contains(text(), "Hóa đơn  có hiệu lực")]'
                )))
                print(" Hóa đơn có hiệu lực (FPT)")
                return "Tìm thấy hóa đơn"
            except TimeoutException:
                print(" Không tìm thấy hóa đơn hoặc hết thời gian")
                return "Không tìm thấy hóa đơn"
        
        elif "meinvoice.vn" in url:
            try:
                wait.until(EC.visibility_of_element_located((By.ID, "popup-content-container")))
                print(" Đã hiển thị kết quả (MISA)")
                return "Tìm thấy hóa đơn"
            except TimeoutException:
                print(" Không tìm thấy hóa đơn (MISA)")
                return "Không tìm thấy hóa đơn"

        elif "van.ehoadon.vn" in url:
            try:
                # Check if invoice exists
                wait.until(EC.presence_of_element_located((By.ID, "frameViewInvoice")))
                print(" Đã tìm thấy hóa đơn (evanhoadon.vn)")
                return "Tìm thấy hóa đơn"
            except TimeoutException:
                print(" Không tìm thấy hóa đơn (evanhoadon.vn)")
                return "Không tìm thấy hóa đơn"

        else:
            print(f" Không hỗ trợ kiểm tra cho trang: {url}")
            return "Không hỗ trợ"

    except Exception as e:
        print(f" Lỗi kiểm tra kết quả: {e}")
        return "Lỗi kiểm tra"

# Hàm tải hóa đơn XML hoặc PDF tùy trang
def tai_hoa_don(driver, url, ma_tra_cuu):
    try:
        if "fpt" in url:
            try:
                wait = WebDriverWait(driver, 10)
                btn = wait.until(EC.element_to_be_clickable((
                    By.XPATH, '//button[contains(text(), "Tải XML")]'
                )))
                driver.execute_script("arguments[0].click();", btn)
                print(" Đã bấm tải XML thành công (FPT)")
                folder = r"D:\RPA\duanFPT"
                
                # Tăng thời gian chờ đợi và kiểm tra kỹ hơn
                if doi_file_tai_xong(folder, timeout=60):  # Tăng timeout lên 60 giây
                    print(" File đã tải xong và được chuyển đổi thành .xml")
                else:
                    print(" File chưa tải xong hoặc lỗi tải")
            except Exception as e:
                 print(f" Lỗi tải XML (FPT): {e}")
            time.sleep(3)

        elif "meinvoice.vn" in url:
            try:
                wait = WebDriverWait(driver, 10)
                folder = r"D:\RPA\duanFPT"
                before_files = set(os.listdir(folder))

                # Mở menu tải hóa đơn
                xpath_menu = '//*[@id="popup-content-container"]/div[1]/div[2]/div[12]/div'
                menu = wait.until(EC.element_to_be_clickable((By.XPATH, xpath_menu)))
                menu.click()
                print(" Đã click menu tải hóa đơn (MISA)")

                # Click nút tải PDF → thực tế là XML
                xpath_pdf = '//*[@id="popup-content-container"]/div[1]/div[2]/div[12]/div/div/div[2]'
                xml_btn = wait.until(EC.element_to_be_clickable((By.XPATH, xpath_pdf)))
                xml_btn.click()
                print(" Đã click nút tải XML (MISA)")

                time.sleep(5)  # Đợi tải file

                # Đổi tên file mới tải về thành ma_tra_cuu.xml
                after_files = set(os.listdir(folder))
                new_file = list(after_files - before_files)
                if new_file:
                    downloaded_file = new_file[0]
                    old_path = os.path.join(folder, downloaded_file)
                    new_name = f"{ma_tra_cuu}.xml"
                    new_path = os.path.join(folder, new_name)
                    os.rename(old_path, new_path)
                    print(f" ✅ Đã đổi tên file MISA thành: {new_name}")
                else:
                    print(" ⚠️ Không tìm thấy file mới tải để đổi tên!")

            except Exception as e:
                print(f" Lỗi khi tải hóa đơn (MISA): {e}")
        elif "van.ehoadon.vn" in url:
            try:
                wait = WebDriverWait(driver, 10)

                # Đợi iframe sẵn sàng và chuyển vào
                wait.until(EC.frame_to_be_available_and_switch_to_it((By.ID, "frameViewInvoice")))
                print(" Đã chuyển vào iframe hóa đơn (evanhoadon.vn)")

                # Đợi nút tải PDF có thể click
                taihoadon = wait.until(EC.element_to_be_clickable((By.ID, "btnDownload")))
                driver.execute_script("arguments[0].click();", taihoadon)
                print(" Đã click nút tải PDF")

                # Đợi nút tải XML có thể click
                taixml = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="LinkDownXML"]')))
                driver.execute_script("arguments[0].click();", taixml)
                print(" Đã click nút tải XML")

                time.sleep(5)

                # Quay lại frame chính
                driver.switch_to.default_content()

            except Exception as e:
                print(f" Lỗi khi tải hóa đơn (evanhoadon.vn): {e}")
    except Exception as e:
        print(f" Lỗi khi tải hóa đơn: {e}")

def trich_xuat_theo_input(df_input, folder_path):
    import xml.etree.ElementTree as ET

    data = []

    def get_text(root, path):
        el = root.find(path)
        return el.text.strip() if el is not None and el.text else ""

    # Lấy danh sách file XML
    xml_files = [f for f in os.listdir(folder_path) if f.endswith(".xml")]

    for i, row in df_input.iterrows():
        mst_input = str(row["Mã số thuế"]).strip()
        ma_tra_cuu = str(row["Mã tra cứu"]).strip()
        url = str(row["URL"]).strip()

        # === Tìm file có chứa mã tra cứu ===
        matched_file = None
        for f in xml_files:
            if ma_tra_cuu.lower() in f.lower():
                matched_file = f
                break

        filepath = os.path.join(folder_path, matched_file) if matched_file else None

        if filepath and os.path.exists(filepath):
            try:
                tree = ET.parse(filepath)
                root = tree.getroot()

                so_hd = get_text(root, ".//TTChung/SHDon")
                don_vi_ban = get_text(root, ".//NBan/Ten")
                mst_ban = get_text(root, ".//NBan/MST")
                dia_chi_ban = get_text(root, ".//NBan/DChi")
                stk_ban = get_text(root, ".//NBan/STKNHang")
                ten_mua = get_text(root, ".//NMua/Ten")
                dia_chi_mua = get_text(root, ".//NMua/DChi")
                mst_mua = get_text(root, ".//NMua/MST")

            except Exception as e:
                print(f" Lỗi đọc file {matched_file}: {e}")
                so_hd = don_vi_ban = mst_ban = dia_chi_ban = stk_ban = ten_mua = dia_chi_mua = mst_mua = ""
        else:
            print(f" Không tìm thấy file XML cho dòng {i+1}: {ma_tra_cuu}")
            so_hd = don_vi_ban = mst_ban = dia_chi_ban = stk_ban = ten_mua = dia_chi_mua = mst_mua = ""

        data.append({
            "STT": i + 1,
            "Mã số thuế": mst_input,
            "Mã tra cứu": ma_tra_cuu,
            "URL": url,
            "Số hóa đơn": so_hd,
            "Đơn vị bán hàng": don_vi_ban,
            "Mã số thuế bên bán": mst_ban,
            "Địa chỉ bên bán": dia_chi_ban,
            "Số tài khoản bên bán": stk_ban,
            "Họ tên người mua hàng": ten_mua,
            "Địa chỉ người mua": dia_chi_mua,
            "Mã số thuế người mua": mst_mua,
        })

    df_out = pd.DataFrame(data)
    output_path = os.path.join(folder_path, "output_hoa_don_final.xlsx")
    df_out.to_excel(output_path, index=False)
    print(f"\n Đã xuất dữ liệu ra: {output_path}")



def main():
    df = pd.read_excel("input.xlsx", dtype={"Mã số thuế": str})
    for index, row in df.iterrows():
        mst = str(row['Mã số thuế']).strip()
        ma_tra_cuu = str(row['Mã tra cứu']).strip()
        url = str(row['URL']).strip()
        if "van.ehoadon.vn" in url and not url.endswith(ma_tra_cuu):
            url += ma_tra_cuu
        print(f"\n Tra cứu dòng {index+1}: {mst} - {ma_tra_cuu} ({url})")

        driver = open_chrome()
        tra_cuu_hoa_don(driver, url, mst, ma_tra_cuu)

        if kiem_tra_ket_qua(driver, url) == "Tìm thấy hóa đơn":
            tai_hoa_don(driver, url,ma_tra_cuu)

        driver.quit()

    trich_xuat_theo_input(df, r"D:\RPA\duanFPT")

main()
