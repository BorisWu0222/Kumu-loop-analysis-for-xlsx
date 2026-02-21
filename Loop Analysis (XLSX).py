import pandas as pd
import networkx as nx
import os

# ================= 路徑與檔名設定 =================
# 自動抓取桌面路徑
desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")

# 輸入檔名
input_filename = 'kumu-boriswu-boriss-intern-mapping-如何解決職涯機構活動參與率低落的問題？-promotion-purpose.xlsx'
input_path = os.path.join(desktop_path, input_filename)

# 輸出檔名 (這會是一個包含詳細報表的 Excel)
output_filename = 'kumu_loops_report_full.xlsx'
output_path = os.path.join(desktop_path, output_filename)

sheet_name = 'Connections' 
# =================================================

def find_loops_and_report():
    # 檢查檔案
    if not os.path.exists(input_path):
        print(f"❌ 錯誤：找不到檔案！請確認桌面是否有：{input_filename}")
        return

    try:
        print(f"📂 正在讀取：{input_path} ...")
        df = pd.read_excel(input_path, sheet_name=sheet_name)
        
        if 'From' not in df.columns or 'To' not in df.columns:
            print("❌ 錯誤：找不到 'From' 或 'To' 欄位。")
            return

        # 建立圖形模型
        G = nx.DiGraph()
        edges = df[['From', 'To']].dropna().values
        G.add_edges_from(edges)

        # 尋找閉環
        print("🔄 正在運算所有閉環 (Complex calculations)...")
        cycles = list(nx.simple_cycles(G))
        total_loops = len(cycles)
        
        if total_loops == 0:
            print("⚠️ 未發現任何閉環。")
            return

        # ==================== 1. VS Code 顯示設定 ====================
        print(f"\n{'='*40}")
        print(f"✅ 運算完成！總共發現 【 {total_loops} 】 個潛在閉環。")
        print(f"{'='*40}\n")

        print("--- 前 5 個閉環範例 (每 3 個變數換行) ---")
        
        for i in range(min(5, total_loops)):
            cycle = cycles[i]
            # 加上起點到最後，形成一個圈
            display_cycle = cycle + [cycle[0]] 
            
            # 格式化輸出字串
            formatted_str = f"Loop {i+1}: "
            indent = " " * len(formatted_str) # 換行後的縮排空格
            
            temp_line = []
            for idx, node in enumerate(display_cycle):
                temp_line.append(node)
                
                # 每 3 個變數，或者已經是最後一個變數時，進行輸出
                if (idx + 1) % 3 == 0 or idx == len(display_cycle) - 1:
                    # 把目前的 temp_line 接起來
                    segment = " -> ".join(temp_line)
                    
                    if idx == len(display_cycle) - 1: # 最後一段
                         # 如果不是該行的第一個元素（也就是接在別人後面），要加箭頭
                        if (idx) % 3 != 0: 
                             formatted_str += " -> " + segment
                        else:
                             formatted_str += "\n" + indent + segment
                    elif idx == 2: # 第一行 (Loop 1: A -> B -> C)
                        formatted_str += segment
                    else: # 中間的行，要換行
                        formatted_str += "\n" + indent + " -> " + segment
                    
                    temp_line = [] # 清空暫存

            print(formatted_str)
            print("-" * 20)

        # ==================== 2. Excel 匯出設定 ====================
        print(f"\n💾 正在產生完整 Excel 報表...")

        # --- 分頁 1: 詳細清單 (Report) ---
        report_data = []
        for i, cycle in enumerate(cycles):
            # 將 list 轉成 "A -> B -> C -> A" 字串
            path_str = " -> ".join(cycle) + " -> " + cycle[0]
            report_data.append({
                'Loop ID': f"Loop {i+1}",
                'Length': len(cycle),
                'Full Path': path_str
            })
        df_report = pd.DataFrame(report_data)

        # --- 分頁 2: Kumu 匯入用 (Import Tags) ---
        # 這是保留給你之後如果要匯回 Kumu 用的
        import_data = []
        for i, cycle in enumerate(cycles):
            loop_tag = f"Loop_{i+1}"
            cycle_edges = list(zip(cycle, cycle[1:] + cycle[:1]))
            for u, v in cycle_edges:
                import_data.append({'From': u, 'To': v, 'Tags': loop_tag})
        
        df_import = pd.DataFrame(import_data)
        # 合併 Tags
        if not df_import.empty:
            df_import = df_import.groupby(['From', 'To'])['Tags'].apply(lambda x: ' | '.join(x)).reset_index()

        # 寫入 Excel (兩個分頁)
        with pd.ExcelWriter(output_path) as writer:
            df_report.to_excel(writer, sheet_name='Loop_Report', index=False)
            df_import.to_excel(writer, sheet_name='For_Kumu_Import', index=False)

        print(f"✅ 成功！檔案已儲存至桌面：{output_filename}")
        print("   - Sheet 1 [Loop_Report]: 包含你要的完整路徑清單。")
        print("   - Sheet 2 [For_Kumu_Import]: 可用來匯入 Kumu 更新標籤。")

    except Exception as e:
        print(f"❌ 發生錯誤: {e}")

if __name__ == "__main__":
    find_loops_and_report()