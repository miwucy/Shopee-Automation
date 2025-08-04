#!/usr/bin/env python3
"""
最簡化的 undetected_chromedriver 測試
"""

import undetected_chromedriver as uc
import time
import os

def minimal_test():
    """最簡化測試"""
    driver = None
    try:
        print("正在啟動 undetected_chromedriver...")
        
        # 最簡化的選項
        options = uc.ChromeOptions()
        options.add_argument('--no-sandbox')
        options.add_argument('--disable-dev-shm-usage')
        
        # 設置用戶數據目錄
        current_dir = os.path.dirname(os.path.abspath(__file__))
        user_data_dir = os.path.join(current_dir, 'Chrome for Testing', 'User Data')
        if os.path.exists(user_data_dir):
            options.add_argument(f'--user-data-dir={user_data_dir}')
            options.add_argument('--profile-directory=Profile 1')
            print(f"使用現有用戶數據目錄: {user_data_dir}")
        
        # 創建驅動器
        driver = uc.Chrome(options=options)
        
        print("正在訪問蝦皮賣家中心...")
        driver.get('https://seller.shopee.tw/portal/sale/shipment?type=toship&source=all&sort_by=create_date_desc')
        
        # 等待頁面加載
        print("等待頁面加載...")
        time.sleep(10)
        
        # 獲取頁面標題
        title = driver.title
        print(f"頁面標題: {title}")
        
        print("✅ 測試完成！")
        print("瀏覽器將保持開啟...")
        
        # 保持瀏覽器開啟
        input("按 Enter 關閉瀏覽器...")
        
    except Exception as e:
        print(f"❌ 測試失敗: {e}")
        import traceback
        traceback.print_exc()
        
    finally:
        if driver:
            driver.quit()

if __name__ == "__main__":
    print("最簡化 undetected_chromedriver 測試")
    print("=" * 30)
    minimal_test() 