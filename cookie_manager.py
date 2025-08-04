
"""
Cookie 管理器 - 用於保存和加載蝦皮登錄狀態
"""

import json
import os
import pickle
from datetime import datetime

class CookieManager:
    def __init__(self, cookie_file="shopee_cookies.pkl"):
        self.cookie_file = cookie_file
    
    def save_cookies(self, driver, domain=".shopee.tw"):
        """保存 cookies 到文件"""
        try:
            cookies = driver.get_cookies()
            
            # 只保存指定域名的 cookies
            filtered_cookies = [cookie for cookie in cookies if domain in cookie.get('domain', '')]
            
            if filtered_cookies:
                with open(self.cookie_file, 'wb') as f:
                    pickle.dump(filtered_cookies, f)
                print(f"✅ 已保存 {len(filtered_cookies)} 個 cookies 到 {self.cookie_file}")
                return True
            else:
                print("⚠️ 未找到相關的 cookies")
                return False
        except Exception as e:
            print(f"❌ 保存 cookies 失敗: {e}")
            return False
    
    def load_cookies(self, driver, domain=".shopee.tw"):
        """從文件加載 cookies"""
        try:
            if not os.path.exists(self.cookie_file):
                print("⚠️ Cookie 文件不存在")
                return False
            
            with open(self.cookie_file, 'rb') as f:
                cookies = pickle.load(f)
            
            # 先訪問目標網站，然後添加 cookies
            driver.get("https://seller.shopee.tw")
            
            success_count = 0
            failed_count = 0
            
            for cookie in cookies:
                try:
                    # 檢查 cookie 的 domain 是否匹配當前頁面
                    cookie_domain = cookie.get('domain', '')
                    current_url = driver.current_url
                    
                    # 如果 domain 不匹配，嘗試修正
                    if cookie_domain and domain not in cookie_domain:
                        # 移除 domain 屬性，讓瀏覽器自動處理
                        cookie_copy = cookie.copy()
                        if 'domain' in cookie_copy:
                            del cookie_copy['domain']
                        driver.add_cookie(cookie_copy)
                    else:
                        driver.add_cookie(cookie)
                    
                    success_count += 1
                    
                except Exception as e:
                    failed_count += 1
                    # 只顯示簡短的錯誤信息，避免冗長的堆疊追蹤
                    error_msg = str(e)
                    if "domain" in error_msg.lower():
                        print(f"⚠️ Cookie domain 不匹配，跳過")
                    else:
                        print(f"⚠️ 添加 cookie 失敗: {error_msg[:100]}...")
                    continue
            
            print(f"✅ 成功加載 {success_count} 個 cookies，失敗 {failed_count} 個")
            
            # 如果至少有一個 cookie 成功加載，就認為是成功的
            return success_count > 0
            
        except Exception as e:
            print(f"❌ 加載 cookies 失敗: {e}")
            return False
    
    def clear_cookies(self):
        """清除保存的 cookies"""
        try:
            if os.path.exists(self.cookie_file):
                os.remove(self.cookie_file)
                print("✅ 已清除 cookies 文件")
                return True
            else:
                print("⚠️ Cookie 文件不存在")
                return False
        except Exception as e:
            print(f"❌ 清除 cookies 失敗: {e}")
            return False
    
    def check_cookies_exist(self):
        """檢查 cookies 文件是否存在"""
        return os.path.exists(self.cookie_file)

def manual_login_helper(driver, cookie_manager):
    """手動登錄輔助函數"""
    print("=" * 50)
    print("🔐 需要手動登錄")
    print("請在瀏覽器中完成登錄，然後按 Enter 繼續...")
    print("=" * 50)
    
    input("登錄完成後按 Enter 繼續...")
    
    # 保存登錄後的 cookies
    if cookie_manager.save_cookies(driver):
        print("✅ 登錄狀態已保存")
    else:
        print("⚠️ 保存登錄狀態失敗")

def auto_login_with_cookies(driver, cookie_manager):
    """使用 cookies 自動登錄"""
    if cookie_manager.check_cookies_exist():
        print("🔄 嘗試使用保存的 cookies 登錄...")
        if cookie_manager.load_cookies(driver):
            # 重新訪問目標頁面
            driver.get('https://seller.shopee.tw/portal/sale/shipment?type=toship&source=all&sort_by=create_date_desc')
            
            # 等待一下讓頁面加載
            import time
            time.sleep(2)
            
            # 檢查是否真的登錄成功（檢查是否有登入按鈕）
            try:
                from selenium.webdriver.common.by import By
                from selenium.webdriver.support.ui import WebDriverWait
                from selenium.webdriver.support import expected_conditions as EC
                from selenium.common.exceptions import TimeoutException
                
                wait = WebDriverWait(driver, 5)
                login_button = wait.until(EC.presence_of_element_located((By.XPATH, "//button[contains(text(), '登入')]")))
                print("⚠️ 檢測到登入按鈕，cookies 可能已過期")
                return False
            except TimeoutException:
                print("✅ Cookies 登錄成功，未檢測到登入按鈕")
                return True
        else:
            print("❌ Cookies 加載失敗")
            return False
    else:
        print("⚠️ 未找到保存的 cookies")
        return False 