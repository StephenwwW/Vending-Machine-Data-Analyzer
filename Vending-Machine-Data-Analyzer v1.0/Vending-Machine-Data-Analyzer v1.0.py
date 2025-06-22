# -*- coding: utf-8 -*-

# 載入必要的系統與第三方函式庫
import sys  # 用於處理系統相關功能，例如結束程式
import re   # 載入正則表達式函式庫，用於進行複雜的文字匹配
import pandas as pd  # 載入 Pandas 函式庫，用於進行高效的資料處理與分析，是本程式的核心
from bs4 import BeautifulSoup  # 載入 BeautifulSoup，用於解析 HTML 原始碼
from PyQt6.QtWidgets import (  # 從 PyQt6 函式庫中載入所有會用到的 GUI 元件
    QApplication, QMainWindow, QWidget, QVBoxLayout, QPushButton, QTableWidget,
    QTableWidgetItem, QLineEdit, QHBoxLayout, QLabel, QHeaderView,
    QMessageBox, QFileDialog, QComboBox, QMenu, QTextEdit, QDialog,
    QDialogButtonBox
)
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QAction

# --- 全域設定 Global Settings ---
# 設定日幣到新台幣的匯率，未來可在此手動更新
JPY_TO_TWD_RATE = 0.20

class SourceViewerDialog(QDialog):
    """
    一個自訂的彈出對話框類別。
    專門用來顯示特定一筆資料的原始HTML碼片段，以供使用者驗證。
    """
    def __init__(self, source_html, parent=None):
        super().__init__(parent)
        self.setWindowTitle("原始碼片段")  # 設定視窗標題
        self.setGeometry(200, 200, 600, 400)  # 設定視窗位置與大小
        
        layout = QVBoxLayout(self) # 使用垂直佈局
        
        # 建立一個文字編輯區來顯示HTML，並設定為唯讀
        text_edit = QTextEdit(self)
        text_edit.setReadOnly(True)
        text_edit.setText(source_html) # 將傳入的HTML文字設定進去
        text_edit.setFontFamily("Courier New") # 使用等寬字體，方便閱讀程式碼
        
        # 建立一個只包含「OK」按鈕的按鈕盒
        button_box = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok)
        button_box.accepted.connect(self.accept) # 將OK按鈕的信號連接到對話框的accept槽(關閉視窗)
        
        # 將元件加入佈局
        layout.addWidget(text_edit)
        layout.addWidget(button_box)

class VendingMachineAnalyzerApp(QMainWindow):
    """
    主應用程式視窗類別。
    包含了所有GUI元件的初始化、功能邏輯以及資料處理流程。
    """
    def __init__(self):
        super().__init__()
        
        # 初始化視窗標題與大小
        self.setWindowTitle('自販機商品分析v1.0')
        self.setGeometry(100, 100, 1000, 700)
        
        # 初始化一個空的 DataFrame，用來存放經過所有處理後的最終資料
        self.processed_df = pd.DataFrame()

        # 建立一個中央的Widget，作為所有GUI元件的容器
        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)

        # 呼叫 UI 初始化函式
        self.initUI()

    def initUI(self):
        """初始化所有使用者介面(GUI)的元件與佈局"""
        
        # 建立主佈局 (垂直)
        main_layout = QVBoxLayout(self.central_widget)
        
        # 建立上半部的控制面板佈局 (水平)
        control_layout = QHBoxLayout()
        
        # 建立核心功能按鈕：「載入原始碼檔案」
        self.parse_button = QPushButton('載入原始碼檔案')
        self.parse_button.setStyleSheet("font-size: 16px; padding: 8px 12px; font-weight: bold;")
        self.parse_button.clicked.connect(self.load_and_process_files) # 連接按鈕點擊事件到主處理函式
        
        # 建立篩選器相關元件
        self.category_filter_label = QLabel('種類篩選:')
        self.category_filter = QComboBox() # 下拉選單用於種類篩選
        self.name_filter_label = QLabel('名稱篩選:')
        self.name_filter = QLineEdit() # 單行輸入框用於名稱篩選
        self.name_filter.setPlaceholderText('輸入關鍵字後按 Enter...')
        self.name_filter.returnPressed.connect(self.apply_filters) # 連接Enter鍵事件到篩選函式

        # 將控制面板的元件依序加入水平佈局
        control_layout.addWidget(self.parse_button)
        control_layout.addStretch() # 加入一個彈簧，讓右側的篩選器靠右對齊
        control_layout.addWidget(self.category_filter_label)
        control_layout.addWidget(self.category_filter)
        control_layout.addWidget(self.name_filter_label)
        control_layout.addWidget(self.name_filter)
        
        # 將控制面板佈局加入主佈局
        main_layout.addLayout(control_layout)

        # 建立顯示資料的表格
        self.table = QTableWidget()
        self.table.verticalHeader().setVisible(False) # 隱藏左側預設的垂直行號
        self.table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers) # 設定表格內容為唯讀，不可編輯
        
        # 設定表格的右鍵選單策略，使其能夠觸發自訂的選單事件
        self.table.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.table.customContextMenuRequested.connect(self.show_context_menu) # 連接右鍵點擊事件到選單顯示函式
        
        main_layout.addWidget(self.table)
        
        # 建立下半部的匯出按鈕佈局 (水平)
        export_layout = QHBoxLayout()
        self.export_csv_button = QPushButton('匯出為 CSV')
        self.export_xlsx_button = QPushButton('匯出為 XLSX')
        self.export_csv_button.clicked.connect(lambda: self.export_data('csv')) # 使用lambda傳遞參數
        self.export_xlsx_button.clicked.connect(lambda: self.export_data('xlsx'))

        export_layout.addStretch() # 加入彈簧使按鈕靠右
        export_layout.addWidget(self.export_csv_button)
        export_layout.addWidget(self.export_xlsx_button)
        main_layout.addLayout(export_layout)
        
        # 連接種類篩選下拉選單的變動事件到篩選函式
        self.category_filter.currentIndexChanged.connect(self.apply_filters)

    def show_context_menu(self, pos):
        """當使用者在表格上按下滑鼠右鍵時，此函式會被觸發"""
        row = self.table.rowAt(pos.y()) # 獲取點擊位置所在的行號
        if row < 0: return # 如果點擊在表格的空白區域，則不執行任何動作

        menu = QMenu() # 建立一個新的右鍵選單
        show_source_action = QAction("顯示原始碼片段", self) # 建立一個名為"顯示原始碼片段"的動作
        # 將這個動作的觸發事件，連接到顯示原始碼的函式，並傳入當前的行號
        show_source_action.triggered.connect(lambda: self.display_source_snippet(row))
        menu.addAction(show_source_action) # 將動作加入選單
        
        # 在滑鼠點擊的位置顯示選單
        menu.exec(self.table.mapToGlobal(pos))

    def display_source_snippet(self, table_row):
        """根據傳入的表格行號，找到對應的資料並彈出視窗顯示其原始碼"""
        try:
            # 獲取使用者點擊的那一行的第0欄(目次)的文字內容
            item_index_str = self.table.item(table_row, 0).text()
            item_index = int(item_index_str)
            
            # 使用該目次，從完整的資料(self.processed_df)中找到對應的原始資料行
            source_row = self.processed_df.loc[self.processed_df['目次'] == item_index]
            
            if not source_row.empty:
                # 提取該行的 '原始碼片段' 欄位內容
                source_html = source_row.iloc[0]['原始碼片段']
                # 建立並顯示我們自訂的原始碼對話框
                dialog = SourceViewerDialog(source_html, self)
                dialog.exec()
        except Exception as e:
            # 如果過程中出錯(例如轉換數字失敗)，則在主控台印出錯誤，避免程式崩潰
            print(f"無法顯示原始碼: {e}")

    def classify_drink(self, name):
        """核心邏輯之一：根據商品名稱中的關鍵字，回傳對應的飲料種類"""
        name = name.lower() # 轉換為小寫以進行不分大小寫的比對
        # 定義分類規則字典，鍵是種類名稱，值是關鍵字列表
        categories = {
            '咖啡': ['コーヒー', 'カフェ', 'ボス', 'ワンダ', 'fire', 'ブラック', 'ラテ', '微糖', 'ブレンド', 'ショット', 'デミタス', 'アロマ'],
            '茶類': ['茶', 'tea', '紅茶', '麦茶', '緑茶', '伊右衛門', '生茶', '颯'],
            '碳酸飲料': ['サイダー', 'ソーダ', 'スカッシュ', 'タンサン', 'coke', 'ペプシ', 'ファンタ', 'メッツ', 'デカビタ'],
            '果汁飲料': ['果実', 'オレンジ', 'りんご', 'ピーチ', 'グレープ', 'レモン', 'ベリー', 'パイン', 'なっちゃん', 'トロピカーナ', 'welch', 'ピングレ', 'うめ'],
            '運動/機能飲料': ['スポーツ', 'ポカリ', 'アミノ', 'dカラ', 'サプリ', 'plus', '免疫', '睡眠', '腸活', 'ラブズ', 'イミューズ', 'カロリミット'],
            '能量飲料': ['モンスター', 'レッドブル', 'zone', 'エナジー', 'ドデカミン'],
            '水': ['水', 'ウォーター', '天然水'],
            '乳製品/其他': ['オレ', 'ミルク', 'ヨーグルト']
        }
        # 遍歷所有種類
        for category, keywords in categories.items():
            # 如果商品名稱包含該種類的任何一個關鍵字
            if any(keyword in name for keyword in keywords):
                return category # 回傳該種類名稱
        return '其他' # 如果都沒匹配到，則歸為「其他」

    def process_data(self, df):
        """核心邏輯之二：對從HTML解析出的原始資料進行完整的二次處理"""
        if df.empty: return pd.DataFrame()
        df_copy = df.copy() # 建立副本以避免修改原始資料

        # 步驟1：清理資料 - 從文字中提取數字並轉換為數值型態
        df_copy['總容量'] = pd.to_numeric(df_copy['商品容量'].str.extract(r'(\d+\.?\d*)')[0], errors='coerce')
        df_copy['日幣售價'] = pd.to_numeric(df_copy['商品價格'].str.extract(r'(\d+\.?\d*)')[0], errors='coerce')
        
        # 步驟2：去除無效資料 - 移除沒有成功提取出容量或價格的行
        df_copy.dropna(subset=['總容量', '日幣售價'], inplace=True)
        df_copy = df_copy[df_copy['總容量'] > 0] # 避免除以零的錯誤

        # 步驟3：進行智慧分類
        df_copy['種類'] = df_copy['商品名稱'].apply(self.classify_drink)

        # 步驟4：進行數學計算
        df_copy['1ml平均價格'] = df_copy['日幣售價'] / df_copy['總容量']
        df_copy['100ml平均價格'] = df_copy['1ml平均價格'] * 100
        df_copy['新台幣售價'] = df_copy['日幣售價'] * JPY_TO_TWD_RATE

        # 步驟5：排序 - 主要按「種類」排序，若種類相同，則按「商品名稱」進行次要排序
        df_copy.sort_values(by=['種類', '商品名稱'], inplace=True)
        
        # 步驟6：建立最終目次 - 在排序後重設索引，以確保目次是連續的
        df_copy.reset_index(drop=True, inplace=True)
        df_copy.insert(0, '目次', df_copy.index + 1)
        
        # 步驟7：整理欄位 - 選取並排列最終要在表格中顯示的欄位
        final_cols = ['目次', '種類', '商品名稱', '總容量', '100ml平均價格', '1ml平均價格', '日幣售價', '新台幣售價', '原始碼片段']
        
        return df_copy[final_cols]

    def load_and_process_files(self):
        """主流程函式：當使用者點擊「載入」按鈕時觸發"""
        # 步驟1：打開檔案選擇對話框，允許多選
        file_paths, _ = QFileDialog.getOpenFileNames(self, "選擇一個或多個原始碼檔案", "", "網頁檔案 (*.html *.htm)")
        if not file_paths: return # 如果使用者取消，則直接返回
        
        all_products, parsed_files_count = [], 0
        
        # 步驟2：遍歷所有選擇的檔案路徑
        for file_path in file_paths:
            try:
                # 讀取檔案內容
                with open(file_path, 'r', encoding='utf-8') as f:
                    html_content = f.read()
                soup = BeautifulSoup(html_content, 'html.parser')

                # 步驟3：智慧偵測檔案格式，並呼叫對應的解析器
                products_from_file = []
                if "okuraya-kanekiya.com" in html_content:
                    products_from_file = self.parse_format_okuraya(soup)
                elif "hachiyoh.co.jp" in html_content:
                    products_from_file = self.parse_format_hachiyoh(soup)
                
                # 如果該檔案成功解析出資料，則加入總列表
                if products_from_file:
                    all_products.extend(products_from_file)
                    parsed_files_count += 1
            except Exception as e:
                # 如果單一檔案讀取或解析失敗，彈出提示並繼續處理下一個檔案
                QMessageBox.warning(self, '檔案讀取錯誤', f'讀取或解析檔案 {file_path} 時發生錯誤:\n{e}')
        
        if not all_products:
            QMessageBox.warning(self, '解析完畢', '所有選擇的檔案均未找到符合已知規則的商品資料。')
            return
        
        # 步驟4：將所有解析出的原始資料轉換為 DataFrame
        raw_df = pd.DataFrame(all_products).rename(columns={'name': '商品名稱', 'capacity': '商品容量', 'price': '商品價格', 'source_html': '原始碼片段'})
        
        # 步驟5：呼叫資料處理函式，進行計算、分類、排序
        self.processed_df = self.process_data(raw_df)
        
        # 步驟6：更新UI介面
        self.update_category_filter() # 更新篩選下拉選單的內容
        self.apply_filters() # 將處理好的資料填入表格
        
        QMessageBox.information(self, '成功', f'資料處理完成！\n成功處理 {parsed_files_count} / {len(file_paths)} 個檔案，共找到 {len(raw_df)} 項商品。')

    def update_category_filter(self):
        """根據讀取到的資料，動態更新「種類篩選」下拉選單中的選項"""
        self.category_filter.blockSignals(True) # 更新前先阻斷信號，避免觸發不必要的事件
        self.category_filter.clear()
        if not self.processed_df.empty:
            # 獲取所有獨一無二的種類名稱，並排序
            categories = ['(全部)'] + sorted(self.processed_df['種類'].unique())
            self.category_filter.addItems(categories)
        self.category_filter.blockSignals(False) # 更新完畢後恢復信號

    def apply_filters(self):
        """根據當前篩選器的設定，過濾DataFrame並更新表格顯示"""
        if self.processed_df.empty:
            self.table.setRowCount(0)
            return
            
        df_to_show = self.processed_df.copy() # 從已處理好的完整資料開始篩選
        
        # 進行種類篩選
        category = self.category_filter.currentText()
        if category and category != '(全部)':
            df_to_show = df_to_show[df_to_show['種類'] == category]
        
        # 進行名稱篩選
        name = self.name_filter.text().strip().lower()
        if name:
            df_to_show = df_to_show[df_to_show['商品名稱'].str.lower().str.contains(name, na=False)]
            
        # 篩選完後，重新產生目次，確保顯示的目次是從1開始的連續數字
        df_to_show.reset_index(drop=True, inplace=True)
        df_to_show['目次'] = df_to_show.index + 1
            
        self.populate_table(df_to_show)

    def populate_table(self, df):
        """將處理好的資料填入GUI表格中"""
        if df.empty:
            self.table.setRowCount(0); self.table.setColumnCount(0)
            return

        headers = ['目次', '種類', '商品名稱', '總容量', '100ml平均價格', '1ml平均價格', '日幣售價', '新台幣售價']
        self.table.setColumnCount(len(headers))
        self.table.setHorizontalHeaderLabels(headers)
        self.table.setRowCount(len(df))

        # 使用 .iterrows() 迭代 DataFrame，並用字典形式取值，可避免因欄位名不符合變數規則而出錯
        for i, row in df.iterrows():
            self.table.setItem(i, 0, QTableWidgetItem(str(row['目次'])))
            self.table.setItem(i, 1, QTableWidgetItem(row['種類']))
            self.table.setItem(i, 2, QTableWidgetItem(row['商品名稱']))
            self.table.setItem(i, 3, QTableWidgetItem(f"{row['總容量']:.0f}"))
            self.table.setItem(i, 4, QTableWidgetItem(f"{row['100ml平均價格']:.2f}"))
            self.table.setItem(i, 5, QTableWidgetItem(f"{row['1ml平均價格']:.4f}"))
            self.table.setItem(i, 6, QTableWidgetItem(f"{row['日幣售價']:.0f}"))
            self.table.setItem(i, 7, QTableWidgetItem(f"{row['新台幣售價']:.0f}"))
        
        self.table.resizeColumnsToContents() # 自動調整欄寬以符合內容

    def parse_format_okuraya(self, soup):
        """專門解析 Okuraya 格式(1.html)的函式"""
        products = []
        lines = soup.find_all('td', class_='line-content')
        for i in range(len(lines)):
            if 'class="drink_content"' in lines[i].get_text():
                # 建立一個字典來存放單一商品的資訊
                current_product = {'name': '', 'capacity': '', 'price': '', 'source_html': ''}
                # 為了驗證功能，將附近的HTML行作為原始碼片段儲存起來
                source_snippet = "\n".join([l.get_text().strip() for l in lines[i:min(i+20, len(lines))]])
                current_product['source_html'] = source_snippet
                # 在接下來的幾行中尋找詳細資訊
                for j in range(i, min(i + 20, len(lines))):
                    line_text = lines[j].get_text()
                    if 'class="drink_title"' in line_text and j + 2 < len(lines):
                        current_product['name'] = lines[j+2].get_text(strip=True).split('<')[0].strip()
                    elif 'class="cost"' in line_text:
                        if j + 2 < len(lines): current_product['capacity'] = lines[j+2].get_text(strip=True).split('<')[0].strip()
                        if j + 4 < len(lines): current_product['price'] = lines[j+4].get_text(strip=True).split('<')[0].strip()
                # 確保名稱、容量、價格都找到了，才算一筆有效資料
                if all(val for key, val in current_product.items() if key != 'source_html'):
                    products.append(current_product)
        return products

    def parse_format_hachiyoh(self, soup):
        """專門解析 Hachiyoh 格式(4.html)的函式"""
        products, prices, names_capacities, sources = [], [], [], []
        lines = soup.find_all('td', class_='line-content')
        # 遍歷所有行，分別收集價格、名稱容量、以及原始碼片段
        for line in lines:
            line_text = line.get_text()
            if 'class="productslist__price"' in line_text:
                match = re.search(r'>([^<]+)<', line_text)
                if match: prices.append(match.group(1).strip())
            elif 'class="productslist__name"' in line_text:
                match = re.search(r'>([^<]+<br>[^<]+)', line_text)
                if match:
                    names_capacities.append(match.group(1).replace('<br>', '|'))
                    sources.append(line.prettify(formatter="html5")) # 儲存該行HTML作為原始碼片段
        # 如果價格和名稱的數量一致，則進行配對組合
        if len(prices) == len(names_capacities) == len(sources):
            for i in range(len(prices)):
                parts = names_capacities[i].split('|')
                products.append({
                    'name': parts[0].strip(),
                    'capacity': parts[1].strip() if len(parts) > 1 else '',
                    'price': prices[i],
                    'source_html': sources[i]
                })
        return products
        
    def export_data(self, file_type):
        """將當前表格中篩選後的資料匯出成 CSV 或 XLSX 檔案"""
        if self.table.rowCount() == 0:
            QMessageBox.warning(self, '警告', '表格中沒有資料可匯出！')
            return
            
        # 重新從 self.processed_df 篩選一次，確保匯出的是篩選後的完整資料
        df_to_export = self.processed_df.copy()
        category = self.category_filter.currentText()
        if category and category != '(全部)': df_to_export = df_to_export[df_to_export['種類'] == category]
        name = self.name_filter.text().strip().lower()
        if name: df_to_export = df_to_export[df_to_export['商品名稱'].str.lower().str.contains(name, na=False)]
        
        # 篩選後重新整理目次以符合匯出
        df_to_export.reset_index(drop=True, inplace=True)
        df_to_export['目次'] = df_to_export.index + 1
        
        # 匯出時排除內部品管用的原始碼片段欄位
        export_cols = ['目次', '種類', '商品名稱', '總容量', '100ml平均價格', '1ml平均價格', '日幣售價', '新台幣售價']
        df_to_export = df_to_export[export_cols]

        filter_str = "CSV 檔案 (*.csv)" if file_type == 'csv' else "Excel 檔案 (*.xlsx)"
        file_path, _ = QFileDialog.getSaveFileName(self, f"儲存報告為 {file_type.upper()}", "", filter_str)
        if file_path:
            try:
                if file_type == 'csv': df_to_export.to_csv(file_path, index=False, encoding='utf-8-sig') 
                elif file_type == 'xlsx': df_to_export.to_excel(file_path, index=False)
                QMessageBox.information(self, '成功', f'報告已成功儲存至:\n{file_path}')
            except Exception as e:
                QMessageBox.critical(self, '儲存錯誤', f'儲存檔案失敗: {e}')

# --- 程式主進入點 ---
if __name__ == '__main__':
    app = QApplication(sys.argv) # 建立應用程式物件
    ex = VendingMachineAnalyzerApp() # 實例化主視窗
    ex.show() # 顯示主視窗
    sys.exit(app.exec()) # 進入應用程式的主事件循環