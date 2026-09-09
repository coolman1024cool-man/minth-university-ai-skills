import os
import sys
import time
import re
import glob
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select

# 設定終端輸出編碼為 UTF-8，防止 Windows CMD (CP950) 報錯閃退
try:
    sys.stdout.reconfigure(encoding='utf-8')
    sys.stderr.reconfigure(encoding='utf-8')
except Exception:
    pass

print("==================================================")
print("   ALLTOP 核銷全自動化機器人 (真·全流程完美版 v7.5)   ")
print("==================================================")

excel_path = '核銷自動化專用表.xlsx'
if not os.path.exists(excel_path):
    print(f"[錯誤] 找不到資料庫表單 {excel_path}！")
    input("按 Enter 鍵結束...")
    sys.exit(1)

try:
    df = pd.read_excel(excel_path).fillna('')
    total_rows = len(df)
    print(f"[成功] 載入 Excel，共 {total_rows} 筆核銷單據。")
except Exception as e:
    print(f"[錯誤] 讀取 Excel 失敗: {e}")
    input("按 Enter 鍵結束...")
    sys.exit(1)

print("\n正在連接至已開啟的 Chrome 瀏覽器 (Port: 9222)...")
chrome_options = Options()
chrome_options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")

try:
    driver = webdriver.Chrome(options=chrome_options)
    print("[成功] 接管 Chrome 瀏覽器！")
except Exception as e:
    print("\n[失敗] 連接 Chrome 失敗！請確認已先執行【1_啟動專屬瀏覽器.bat】。")
    print(f"錯誤細節: {e}")
    input("\n按 Enter 鍵結束...")
    sys.exit(1)

def switch_to_main():
    driver.switch_to.default_content()
    try:
        driver.switch_to.frame("iframe")
        return True
    except:
        iframes = driver.find_elements(By.TAG_NAME, "iframe")
        for f in iframes:
            if f.get_attribute("id") == "iframe":
                driver.switch_to.frame(f)
                return True
    return False

def select_payee_natively(payee_code):
    """點擊藍色選擇按鈕，並在彈窗 iframe_windows 中原生搜尋並選定受款人"""
    print(f"  [查詢] 正在透過學校資料庫原生查詢代號: 【{payee_code}】...")
    switch_to_main()
    
    sel_btns = driver.find_elements(By.XPATH, "//button[contains(@onclick, 'SelectTbpay.php')]")
    if not sel_btns:
        print("  [提示] 找不到受款人 [選擇] 按鈕")
        return False
    driver.execute_script("arguments[0].click();", sel_btns[0])
    time.sleep(2)
    
    driver.switch_to.default_content()
    try:
        driver.switch_to.frame("iframe_windows")
    except:
        print("  [警告] 無法切換至查詢彈窗 iframe_windows")
        switch_to_main()
        return False
        
    search_inp = driver.find_element(By.ID, "key_srh")
    search_inp.clear()
    search_inp.send_keys(str(payee_code))
    
    search_btn = driver.find_element(By.CSS_SELECTOR, "button.search")
    driver.execute_script("arguments[0].click();", search_btn)
    time.sleep(2)
    
    green_btns = driver.find_elements(By.CSS_SELECTOR, "button.green")
    if green_btns:
        driver.execute_script("arguments[0].click();", green_btns[0])
        time.sleep(2)
        print(f"  [成功] 受款人原生選定，已取得伺服器簽發之加密金鑰！")
    else:
        print(f"  [警告] 查詢不到代號 【{payee_code}】")
        
    switch_to_main()
    return True

def fill_invoice_modal_natively(inv_type, tax_id, date_str, amount, inv_no, item_name):
    """自動點開發票明細彈窗，並填寫憑證資料後確認存檔"""
    print(f"  [憑證] 正在自動填報發票/憑證明細: 【{inv_type}】 統編: {tax_id} 金額: ${amount}...")
    switch_to_main()
    
    # 點擊發票明細的新增按鈕
    inv_add_btns = driver.find_elements(By.XPATH, "//button[contains(text(), '新增') and contains(@class, 'btn-info')] | //button[contains(@onclick, 'formPay_81.php')]")
    if not inv_add_btns:
        print("  [警告] 找不到新增發票明細按鈕")
        return False
        
    driver.execute_script("arguments[0].click();", inv_add_btns[-1])
    time.sleep(2)
        
    driver.switch_to.default_content()
    try:
        driver.switch_to.frame("iframe_windows")
    except:
        print("  [警告] 無法切換至發票彈窗 iframe_windows")
        switch_to_main()
        return False
        
    # 日期西元格式 (例如 115/08/20 -> 2026-08-20)
    iso_date = str(date_str).replace('/', '-')
    if len(iso_date) >= 7 and iso_date.startswith("115"):
        iso_date = "2026" + iso_date[3:]
        
    # 判斷是否為標準 2 英文 + 8 數字發票號碼
    inv_no_clean = str(inv_no).strip().upper()
    is_valid_inv_no = bool(re.match(r'^[A-Z]{2}\d{8}$', inv_no_clean))
    
    if "發票" in inv_type and "免用" not in inv_type and is_valid_inv_no:
        # 標準發票模式 (需要 2 英文 + 8 數字)
        print(f"  [類別] 選取【發票】模式，號碼: {inv_no_clean}")
        driver.execute_script(f"""
            $('#billnum').val('{inv_no_clean}').change();
            $('#payno').val('{tax_id}').change();
            $('#billdate').val('{iso_date}').change();
            $('#money').val('{amount}').change();
        """)
    else:
        # 免用統一發票 / 收據模式 (品名欄位為 goodsname)
        print(f"  [類別] 選取【免用統一發票】模式，品名: {item_name}")
        driver.execute_script("""
            $('#billnumkind').val('免用統一發票').change();
            document.getElementById('form-main').submit();
        """)
        time.sleep(2)
        
        driver.switch_to.default_content()
        driver.switch_to.frame("iframe_windows")
        
        driver.execute_script(f"""
            if (document.getElementById('goodsname')) {{
                document.getElementById('goodsname').value = '{item_name}';
                $('#goodsname').change();
            }}
            document.getElementById('payno').value = '{tax_id}';
            document.getElementById('billdate').value = '{iso_date}';
            document.getElementById('money').value = '{amount}';
            $('#payno').change();
            $('#billdate').change();
            $('#money').change();
        """)
        
    time.sleep(1)
    
    # 點擊確認按鈕存檔
    save_btn = driver.find_elements(By.CSS_SELECTOR, "button.save, button.btn-success")
    if save_btn:
        driver.execute_script("arguments[0].click();", save_btn[0])
        time.sleep(2.5)
        print("  [成功] 發票明細建立成功，彈窗已自動儲存關閉！")
        
    switch_to_main()
    return True

switch_to_main()

# =========================================================================
# 階段 1: 處理主單用途說明與儲存 (若當前在 formPay_1.php?cmd=new)
# =========================================================================
current_url = driver.execute_script("return window.location.href;")
print(f"當前工作頁面: {current_url}")

if "formPay_1.php" in current_url:
    print("\n【階段 1】自動填寫主單「用途說明」...")
    first_row = df.iloc[0]
    purpose = str(first_row.get('用途說明', '')).strip()
    
    memo_field = driver.find_elements(By.ID, "differencememo")
    if memo_field and memo_field[0].is_displayed():
        curr_val = memo_field[0].get_attribute("value")
        if not curr_val:
            memo_field[0].clear()
            memo_field[0].send_keys(purpose)
            time.sleep(0.8)
            print(f"  [填入] 用途說明: 【{purpose}】")
            
            # 點擊確認存檔主單
            save_btn = driver.find_elements(By.CSS_SELECTOR, "button.save, button.btn-success")
            if save_btn:
                print("  [存檔] 正在儲存主單...")
                driver.execute_script("arguments[0].click();", save_btn[0])
                time.sleep(3)
                switch_to_main()

    # 階段 1.2: 新增「核銷明細」科目
    print("\n【階段 1.2】自動建立「核銷明細」品名與金額...")
    switch_to_main()
    
    for idx, row in df.iterrows():
        item_name = str(row.get('使用說明(品名)', '')).strip()
        amount = str(row.get('申請金額', '')).strip()
        bud_code = str(row.get('預算編號', '')).strip()
        
        page_src = driver.page_source
        if item_name in page_src and str(amount) in page_src:
            print(f"  [提示] 科目 【{item_name}】 ${amount} 已存在，略過新增。")
            continue
            
        print(f"  [新增] 正在新增科目: 【{item_name}】 ${amount}...")
        add_btns = driver.find_elements(By.XPATH, "//button[contains(text(), '新增') and contains(@class, 'btn-info')]")
        if add_btns:
            driver.execute_script("arguments[0].click();", add_btns[0])
            time.sleep(2)
            switch_to_main()
            
            # 點擊選擇預算來源
            sel_bud_btn = driver.find_elements(By.XPATH, "//button[contains(@onclick, 'SelectBudget.php')]")
            if sel_bud_btn:
                driver.execute_script("arguments[0].click();", sel_bud_btn[0])
                time.sleep(2)
                driver.switch_to.default_content()
                driver.switch_to.frame("iframe_windows")
                
                # 搜尋或選定預算項目
                found = False
                for p_num in range(1, 4):
                    for tr in driver.find_elements(By.TAG_NAME, "tr"):
                        if bud_code in tr.text or "業務費" in tr.text:
                            b = tr.find_elements(By.CSS_SELECTOR, "button.green, button.btn-success")
                            if b:
                                driver.execute_script("arguments[0].click();", b[0])
                                found = True
                                break
                    if found: break
                    # 翻頁
                    next_links = driver.find_elements(By.XPATH, f"//a[text()='{p_num+1}']")
                    if next_links:
                        driver.execute_script("arguments[0].click();", next_links[0])
                        time.sleep(2)
                    else: break
                time.sleep(2)
                switch_to_main()
            
            r_inp = driver.find_elements(By.ID, "REMARK")
            if r_inp: r_inp[0].send_keys(item_name)
            
            m_inp = driver.find_elements(By.ID, "money")
            if m_inp: m_inp[0].send_keys(amount)
            
            time.sleep(0.5)
            save_b = driver.find_elements(By.CSS_SELECTOR, "button.save, button.btn-success")
            if save_b:
                driver.execute_script("arguments[0].click();", save_b[-1])
                time.sleep(2.5)
                switch_to_main()
                print(f"  [成功] 科目 【{item_name}】 ${amount} 已儲存！")

    # 切換至付款明細分頁
    print("\n【階段 1.3】切換至「付款明細」分頁...")
    pay_tabs = driver.find_elements(By.XPATH, "//a[contains(text(), '付款明細')]")
    if pay_tabs:
        driver.execute_script("arguments[0].click();", pay_tabs[0])
        time.sleep(2)
        switch_to_main()

# =========================================================================
# 階段 2: 逐筆處理付款明細與發票明細
# =========================================================================
print("\n【階段 2】逐筆自動建立「付款明細」與「發票憑證」...")
switch_to_main()

for idx, row in df.iterrows():
    item_num = idx + 1
    item_name = str(row.get('使用說明(品名)', '')).strip()
    amount = str(row.get('申請金額', '')).strip()
    pay_type = str(row.get('付款類別', '付款資料')).strip()
    payee_id = str(row.get('受款人代號', '')).strip()
    payee_name = str(row.get('受款人姓名', '')).strip()
    id_type = str(row.get('付款對象身分', '廠商')).strip()
    income_type = str(row.get('所得類別', '0-不扣所得')).strip()
    inv_type = str(row.get('憑證類別', '發票')).strip()
    tax_id = str(row.get('廠商統編', '')).strip()
    date_str = str(row.get('單據日期', '')).strip()
    inv_no = str(row.get('發票號碼', '')).strip()
    
    current_url = driver.execute_script("return window.location.href;")
    
    # 若在付款清單頁面 (formPay_6.php)，點擊新增
    if "formPay_6.php" in current_url:
        print(f"\n[第 {item_num}/{total_rows} 筆] 點擊【新增】進入付款明細編輯...")
        new_btns = driver.find_elements(By.CSS_SELECTOR, "button.new, button.btn-info")
        for nb in new_btns:
            if "新增" in nb.text:
                driver.execute_script("arguments[0].click();", nb)
                time.sleep(2)
                switch_to_main()
                break

    # 若已在付款明細編輯頁面 (formPay_61.php)
    current_url = driver.execute_script("return window.location.href;")
    if "formPay_61.php" in current_url:
        print(f"\n[填寫] 正在全自動填寫第 {item_num} 筆: 【{item_name}】 ${amount} (受款人: {payee_name})")
        
        # 設置付款類別 (付款資料 val=1, 所得資料 val=2, 付款兼所得 val=3)
        ispay_val = "1"
        if "付款資料" in pay_type: ispay_val = "1"
        elif pay_type == "所得資料": ispay_val = "2"
        elif "付款兼所得" in pay_type or "代墊" in pay_type: ispay_val = "3"
        
        # 設置付款身分 (廠商 val=1, 學生 val=2, 教職員 val=3, 校外 val=4)
        paybus_val = "1" if "廠商" in id_type else ("2" if "學生" in id_type else ("4" if "校外" in id_type else "3"))
        
        driver.execute_script(f"""
            if(typeof $.fn.selectpicker !== 'undefined'){{
                $('#IsPay').selectpicker('val', '{ispay_val}');
                $('#salkind').selectpicker('val', '0');
                $('#PayBusKind').selectpicker('val', '{paybus_val}');
            }}
            $('#IsPay').val('{ispay_val}').change();
            $('#salkind').val('0').change();
            $('#PayBusKind').val('{paybus_val}').change();
        """)
        time.sleep(1)

        # 選取主單對應科目
        try:
            formid_sel = Select(driver.find_element(By.ID, "formid"))
            for opt in formid_sel.options:
                if str(amount) in opt.text or item_name in opt.text:
                    opt_val = opt.get_attribute("value")
                    formid_sel.select_by_value(opt_val)
                    driver.execute_script(f"$('#formid').selectpicker('val', '{opt_val}').change();")
                    time.sleep(0.8)
                    print(f"  [科目] 已選取主單科目: {opt.text}")
                    break
        except Exception as e:
            print(f"  [警告] 選取科目: {e}")

        # 呼叫原生彈窗查詢受款人 (自動綁定資安金鑰)
        select_payee_natively(payee_id)

        # 填寫費用說明與金額
        driver.execute_script(f"""
            $('#salitem').val('{item_name}').change();
            $('#salmoney').val('{amount}').change();
            document.getElementById('money').value = '{amount}';
        """)
        time.sleep(0.8)
        
        # 先儲存付款明細表頭 (進入編輯模式以便開啟發票明細)
        save_btn = driver.find_elements(By.CSS_SELECTOR, "button.save, button.btn-success")
        if save_btn:
            driver.execute_script("arguments[0].click();", save_btn[0])
            time.sleep(2.5)
            switch_to_main()

        # 全自動打開發票彈窗並填入確認！
        fill_invoice_modal_natively(inv_type, tax_id, date_str, amount, inv_no, item_name)
        
        # 再次確認存檔此筆付款明細並返回清單
        save_btn = driver.find_elements(By.CSS_SELECTOR, "button.save, button.btn-success")
        if save_btn:
            driver.execute_script("arguments[0].click();", save_btn[-1])
            time.sleep(2.5)
            switch_to_main()
            print(f"  [完成] 第 {item_num} 筆付款明細已完整建立並儲存！")

# =========================================================================
# 階段 3: 自動上傳附件 (若有指定附件資料夾)
# =========================================================================
print("\n【階段 3】自動掃描並上傳單據附件...")
first_row = df.iloc[0]
att_dir = str(first_row.get('附件資料夾路徑', '')).strip()

if att_dir and os.path.exists(att_dir):
    switch_to_main()
    att_tabs = driver.find_elements(By.XPATH, "//a[contains(text(), '附件')]")
    if att_tabs:
        driver.execute_script("arguments[0].click();", att_tabs[0])
        time.sleep(2)
        switch_to_main()
        
        # 尋找圖檔與 PDF
        files = glob.glob(os.path.join(att_dir, "*.jpg")) + glob.glob(os.path.join(att_dir, "*.png")) + glob.glob(os.path.join(att_dir, "*.pdf"))
        print(f"  [附件] 在資料夾找到 {len(files)} 個檔案準備上傳...")
        
        for f_path in files:
            add_att_btn = driver.find_elements(By.CSS_SELECTOR, "button.new")
            if add_att_btn:
                driver.execute_script("arguments[0].click();", add_att_btn[0])
                time.sleep(1)
                
            file_inps = driver.find_elements(By.XPATH, "//input[@type='file']")
            if file_inps:
                file_inps[-1].send_keys(f_path)
                print(f"  [上傳] 已掛載附件: {os.path.basename(f_path)}")
                time.sleep(1)
                
        # 點擊確認存檔附件
        save_att_btn = driver.find_elements(By.CSS_SELECTOR, "button.save")
        if save_att_btn:
            driver.execute_script("arguments[0].click();", save_att_btn[0])
            time.sleep(3)
            print("  [成功] 全部附件已成功上傳並儲存！")

print("\n" + "="*50)
print("  🎉 恭喜！整張核銷單已 100% 全自動建立完成！")
print("==================================================")
