import os
import json
import sys

# 模擬 AppConfig 和 FileManager 的核心邏輯
class AppConfig:
    def __init__(self):
        self.base_path = os.path.dirname(os.path.abspath(__file__))
        if getattr(sys, "frozen", False):
            self.base_path = os.path.dirname(sys.executable)
            
        self.config_filename = "lmreview_config.json"
        self.projects = ["【雲端案】", "【整合案】", "【Trod案】"]
        self.deliveries = ["【契約交付】", "【其他交付】"]
        self.input_folder = "input"
        
        self._load_config()

    def _load_config(self):
        path = os.path.join(self.base_path, self.config_filename)
        if os.path.exists(path):
            try:
                with open(path, "r", encoding="utf-8") as fh:
                    data = json.load(fh)
                    print(f"[DEBUG] 讀取到設定檔: {path}")
                    if data.get("projects"): self.projects = data["projects"]
                    if data.get("deliveries"): self.deliveries = data["deliveries"]
            except Exception as e:
                print(f"[ERROR] 設定檔讀取錯誤: {e}")
        else:
            print(f"[DEBUG] 無設定檔，使用預設值。路徑: {path}")

def list_files(cfg):
    print(f"\n[INFO] 程式基底路徑: {cfg.base_path}")
    
    for p in cfg.projects:
        for d in cfg.deliveries:
            target_dir = os.path.join(cfg.base_path, p, d, cfg.input_folder)
            print(f"\n檢查目錄: {target_dir}")
            
            if not os.path.exists(target_dir):
                print("  -> ❌ 目錄不存在")
                continue
                
            try:
                files = os.listdir(target_dir)
                if not files:
                    print("  -> ⚠️ 目錄是空的")
                else:
                    print(f"  -> ✅ 發現 {len(files)} 個檔案:")
                    for f in files:
                        print(f"     - {f}")
            except Exception as e:
                print(f"  -> ❌ 無法讀取: {e}")

if __name__ == "__main__":
    try:
        cfg = AppConfig()
        list_files(cfg)
        input("\n按 Enter 鍵結束...")
    except Exception as e:
        print(f"發生未預期的錯誤: {e}")
        input("按 Enter 鍵結束...")
