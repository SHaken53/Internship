import os
import sys
import csv
import itertools
import pandas as pd
from openpyxl import Workbook, load_workbook
from PyQt5 import QtWidgets, QtGui
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QPixmap
from PyQt5 import QtCore
from PyQt5.QtWidgets import QApplication, QMainWindow, QFileDialog, QVBoxLayout, QHBoxLayout, \
    QLabel, QLineEdit, QPushButton, QComboBox, QTextEdit, QMessageBox, QWidget, \
    QListWidget, QListWidgetItem, QTableWidget, QTableWidgetItem, QTabWidget, QDialog, \
    QFrame, QSizePolicy, QHeaderView, QCheckBox



class ExcelProcessorApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.init_ui()
        
        self.column_checkboxes = []  # Sütun checkboxları için boş bir liste
        self.selected_columns = []  # Seçili sütunların indekslerini tutacak bir liste
        self.column_names = ['Sütun 1', 'Sütun 2', 'Sütun 3']  # Sütun adlarını içeren bir liste

        self.create_column_checkboxes()
        

    def init_ui(self):
        self.setWindowTitle("Format_Düzenleyici_Demo_01")
        self.setGeometry(100, 100, 1920, 1080)
        self.showMaximized()

        main_widget = QWidget(self)
        self.setCentralWidget(main_widget)

        layout = QVBoxLayout()
        main_widget.setLayout(layout)

        # Yeni bir QVBoxLayout oluştur ve tüm sonuçları bu düzenleyiciye yerleştir
        self.results_layout = QVBoxLayout()
        self.results_layout.setAlignment(Qt.AlignTop)
        layout.addLayout(self.results_layout)

        button_style = """
            QPushButton {
                border: 2px solid #4085FF;
                border-radius: 16px;
                background-color: #4085FF;
                color: white;
                font-size: 15px;
                padding: 10px 20px;
            }
            QPushButton:hover {
                background-color: #2e5599;
            }
            QPushButton:pressed {
                background-color: #2e5599;
            }
        """

        logo_label = QLabel(self)
        pixmap = QPixmap("logo.png")
        scaled_pixmap = pixmap.scaled(100, 100, Qt.KeepAspectRatio, Qt.SmoothTransformation)
        logo_label.setPixmap(scaled_pixmap)
        layout.addWidget(logo_label)

        title_label_style = "background-color: #0030B4; color: white; font-size: 18px; font-weight: bold; padding: 5px;"

        self.file_path_label = QLabel("Excel Dosyası Seçin:")
        self.file_path_label.setStyleSheet(title_label_style)
        layout.addWidget(self.file_path_label)

        self.file_path_input = QLineEdit()
        layout.addWidget(self.file_path_input)

        # Browse ve Convert butonlarının olduğu layout
        browse_and_convert_layout = QHBoxLayout()
        browse_and_convert_layout.addStretch(1)
        self.browse_button = QPushButton("Gözat")
        self.browse_button.clicked.connect(self.browse_file)
        self.browse_button.setFixedWidth(200)
        self.browse_button.setFixedHeight(55)
        self.browse_button.setStyleSheet(button_style)
        browse_and_convert_layout.addWidget(self.browse_button)

        # Düz dikey çizgi ekleniyor
        line = QFrame(self)
        line.setFrameShape(QFrame.VLine)
        line.setFixedHeight(40)
        line.setStyleSheet("color: gray;")
        browse_and_convert_layout.addWidget(line)

        self.convert_button = QPushButton("Excel'i CSV'ye Dönüştür")
        self.convert_button.clicked.connect(self.convert_to_csv)
        self.convert_button.setFixedWidth(200)
        self.convert_button.setFixedHeight(55)
        self.convert_button.setStyleSheet(button_style)
        browse_and_convert_layout.addWidget(self.convert_button)
        browse_and_convert_layout.addStretch(1)
        layout.addLayout(browse_and_convert_layout)

        # Düz dikey çizgi ekleniyor
        line = QFrame(self)
        line.setFrameShape(QFrame.HLine)
        line.setStyleSheet("color: gray;")
        layout.addWidget(line)

        self.column_label = QLabel("İşlem yapılacak sütunları seçin:")
        self.column_label.setStyleSheet(title_label_style)
        layout.addWidget(self.column_label)

        self.column_list = QListWidget()
        layout.addWidget(self.column_list)
        # Sütun adlarının stilini güncelle
        column_style = "font-size: 16px; padding: 10px 0;"
        # Sütun adlarını daha büyük hale getir
        font = self.column_list.font()
        font.setPointSize(16)  # İstediğiniz yazı tipi boyutunu burada ayarlayabilirsiniz
        self.column_list.setFont(font)

        # Sütun adlarını daha büyük boyutlu hale getir
        for i in range(self.column_list.count()):
            item = self.column_list.item(i)
            item.setTextAlignment(Qt.AlignCenter)
            item.setText(item.text().upper())  # Sütun adını büyük harfle göster
            item.setStyleSheet(column_style)

        # Tümünü Seç ve Tümünden Vazgeç butonlarının olduğu layout
        button_row_layout = QHBoxLayout()
        button_row_layout.addStretch(1)
        self.select_all_button = QPushButton("Tümünü Seç")
        self.select_all_button.clicked.connect(self.select_all_columns)
        self.select_all_button.setFixedWidth(175)
        self.select_all_button.setFixedHeight(55)
        self.select_all_button.setStyleSheet(button_style)
        button_row_layout.addWidget(self.select_all_button)

        # Düz dikey çizgi ekleniyor
        line = QFrame(self)
        line.setFrameShape(QFrame.VLine)
        line.setFixedHeight(40)
        line.setStyleSheet("color: gray;")
        button_row_layout.addWidget(line)

        self.deselect_all_button = QPushButton("Tümünden Vazgeç")
        self.deselect_all_button.clicked.connect(self.deselect_all_columns)
        self.deselect_all_button.setFixedWidth(175)
        self.deselect_all_button.setFixedHeight(55)
        self.deselect_all_button.setStyleSheet(button_style)
        button_row_layout.addWidget(self.deselect_all_button)
        button_row_layout.addStretch(1)
        layout.addLayout(button_row_layout)

        # VARYASYONKODU ve diğerleri için layout
        operation_buttons_layout = QHBoxLayout()
        operation_buttons_layout.addStretch(1)
        
        self.variation_kod_button = QPushButton("VARYASYONKODU")
        self.variation_kod_button.clicked.connect(self.process_variation_kod)
        self.variation_kod_button.setFixedWidth(200)
        self.variation_kod_button.setFixedHeight(55)
        self.variation_kod_button.setStyleSheet(button_style)
        operation_buttons_layout.addWidget(self.variation_kod_button)

        # Düz dikey çizgi ekleniyor
        line = QFrame(self)
        line.setFrameShape(QFrame.VLine)
        line.setFixedHeight(40)
        line.setStyleSheet("color: gray;")
        operation_buttons_layout.addWidget(line)

        self.breadcrumb_kat_button = QPushButton("BREADCRUMBKAT")
        self.breadcrumb_kat_button.clicked.connect(self.process_breadcrumb_kat)
        self.breadcrumb_kat_button.setFixedWidth(200)
        self.breadcrumb_kat_button.setFixedHeight(55)
        self.breadcrumb_kat_button.setStyleSheet(button_style)
        operation_buttons_layout.addWidget(self.breadcrumb_kat_button)

        # Düz dikey çizgi ekleniyor
        line = QFrame(self)
        line.setFrameShape(QFrame.VLine)
        line.setFixedHeight(40)
        line.setStyleSheet("color: gray;")
        operation_buttons_layout.addWidget(line)

        self.variation_button = QPushButton("VARYASYON")
        self.variation_button.clicked.connect(self.process_variation)
        self.variation_button.setFixedWidth(200)
        self.variation_button.setFixedHeight(55)
        self.variation_button.setStyleSheet(button_style)
        operation_buttons_layout.addWidget(self.variation_button)

        # Düz dikey çizgi ekleniyor
        line = QFrame(self)
        line.setFrameShape(QFrame.VLine)
        line.setFixedHeight(40)
        line.setStyleSheet("color: gray;")
        operation_buttons_layout.addWidget(line)

        # "Kategoriler" düğmesini "VARYASYON" düğmesinin hemen sağına ekliyoruz
        self.kategoriler_dugmesi_1 = QPushButton("KATEGORİLER-1")
        self.kategoriler_dugmesi_1.clicked.connect(self.process_categories_1)
        self.kategoriler_dugmesi_1.setFixedWidth(200)
        self.kategoriler_dugmesi_1.setFixedHeight(55)
        self.kategoriler_dugmesi_1.setStyleSheet(button_style)
        operation_buttons_layout.addWidget(self.kategoriler_dugmesi_1)

        # Düz dikey çizgi ekleniyor
        line = QFrame(self)
        line.setFrameShape(QFrame.VLine)
        line.setFixedHeight(40)
        line.setStyleSheet("color: gray;")
        operation_buttons_layout.addWidget(line)

        # "Kategoriler" düğmesini "VARYASYON" düğmesinin hemen sağına ekliyoruz
        self.kategoriler_dugmesi_2 = QPushButton("KATEGORİLER-2")
        self.kategoriler_dugmesi_2.clicked.connect(self.process_categories_2)
        self.kategoriler_dugmesi_2.setFixedWidth(200)
        self.kategoriler_dugmesi_2.setFixedHeight(55)
        self.kategoriler_dugmesi_2.setStyleSheet(button_style)
        operation_buttons_layout.addWidget(self.kategoriler_dugmesi_2)

        operation_buttons_layout.addStretch(1)
        layout.addLayout(operation_buttons_layout)


        self.operations_label = QLabel("İşlemleri burada görüntüleyin:")
        self.operations_label.setStyleSheet(title_label_style)
        layout.addWidget(self.operations_label)

        self.operations_tabs = QTabWidget()
        layout.addWidget(self.operations_tabs)

        # Kaydet ve Tüm İşlemlerden Vazgeç butonları için layout
        operation_buttons_layout2 = QHBoxLayout()
        operation_buttons_layout2.addStretch(1)
        self.save_button = QPushButton("Kaydet")
        self.save_button.clicked.connect(self.save_results_to_excel)
        self.save_button.setFixedWidth(200)
        self.save_button.setFixedHeight(55)
        self.save_button.setStyleSheet(button_style)
        operation_buttons_layout2.addWidget(self.save_button)

        # Düz dikey çizgi ekleniyor
        line = QFrame(self)
        line.setFrameShape(QFrame.VLine)
        line.setFixedHeight(40)
        line.setStyleSheet("color: gray;")
        operation_buttons_layout2.addWidget(line)

        self.cancel_button = QPushButton("Tüm İşlemlerden Vazgeç")
        self.cancel_button.clicked.connect(self.cancel_operations)
        self.cancel_button.setFixedWidth(200)
        self.cancel_button.setFixedHeight(55)
        self.cancel_button.setStyleSheet(button_style)
        operation_buttons_layout2.addWidget(self.cancel_button)
        operation_buttons_layout2.addStretch(1)
        layout.addLayout(operation_buttons_layout2)

        self.file_path = ""
        self.data_frame = None
        self.csv_file_path = ""

        self.variation_kod_results = []
        self.breadcrumb_kat_results = []
        self.variation_results = []
        self.categories_results = []

        self.load_columns()

    def browse_file(self):
        options = QFileDialog.Options()
        file_path, _ = QFileDialog.getOpenFileName(self, "Excel Dosyası Seçin", "", "Excel Dosyaları (*.xlsx *.xls);;All Files (*)", options=options)

        if file_path:
            self.file_path = file_path
            self.file_path_input.setText(file_path)
            self.load_columns()

    def convert_to_csv(self):
        if not self.file_path:
            QMessageBox.warning(self, "Uyarı", "Lütfen önce bir Excel dosyası seçin.")
            return

        try:
            # Excel dosyasını oku ve CSV dosyasına dönüştür
            excel_data = pd.read_excel(self.file_path)
            self.csv_file_path = self.file_path.replace(".xlsx", ".csv").replace(".xls", ".csv")
            excel_data.to_csv(self.csv_file_path, index=False)
            QMessageBox.information(self, "Bilgi", "Excel dosyası başarıyla CSV'ye dönüştürüldü.")
            self.file_path_input.setText(self.csv_file_path)
            self.load_columns()

            # Tüm sütunları seçimden kaldır
            for i in range(self.column_list.count()):
                item = self.column_list.item(i)
                item.setCheckState(Qt.Unchecked)

        except Exception as e:
            QMessageBox.critical(self, "Hata", "Excel dosyası dönüştürülürken bir hata oluştu:\n" + str(e))

    def load_columns(self):
        if not self.csv_file_path:
            return
        try:
            # Dosyanın gerçek kodlamasını tespit etme
            with open(self.csv_file_path, 'rb') as f:
                raw_data = f.read()

            # Olası kodlamaları deneyerek dosyayı açma
            for encoding in ["utf-8-sig", "latin-1", "cp1254"]:
                try:
                    self.data_frame = pd.read_csv(self.csv_file_path, encoding=encoding)
                    break
                except UnicodeDecodeError:
                    continue
            else:
                # Hiçbir kodlama ile başarılı açma işlemi yapılamadıysa hata ver
                QMessageBox.critical(self, "Hata", "Dosya uygun bir kodlama ile açılamadı.")
                return

        except Exception as e:
            QMessageBox.critical(self, "Hata", "Dosya açılırken bir hata oluştu:\n" + str(e))
            return

        columns = self.data_frame.columns

        # Sütun başlıklarını ve ilk 5 veriyi göstermek için tablo oluşturma
        table_widget = QTableWidget(6, len(columns))
        table_widget.setHorizontalHeaderLabels(columns)

        for c, column in enumerate(columns):
            values = self.data_frame[column].head(5).tolist()
            for r, value in enumerate(values):
                item = QTableWidgetItem(str(value))
                table_widget.setItem(r, c, item)

        # Sütun adlarını da ilk satıra ekleme
        for c, column in enumerate(columns):
            item = QTableWidgetItem(column)
            table_widget.setItem(5, c, item)

        # Tabloyu bir öğe olarak liste içine yerleştirmek için widget kullanma
        item = QListWidgetItem()
        item.setSizeHint(table_widget.sizeHint())

        self.column_list.addItem(item)
        self.column_list.setItemWidget(item, table_widget)

        # Sütun adlarına tıklama özelliğini ekle
        self.column_list.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.column_list.itemClicked.connect(self.toggle_column_selection)
    
    def create_column_checkboxes(self):
        for i, column_name in enumerate(self.column_names):
            checkbox = QCheckBox(column_name)
            checkbox.stateChanged.connect(lambda state, idx=i: self.toggle_column_selection(state, idx))
            self.column_checkboxes.append(checkbox)
          
    def toggle_column_selection(self, state, index):
        checkbox = self.column_checkboxes[index]
        if state == QtCore.Qt.Checked:
            if index not in self.selected_columns:
                self.selected_columns.append(index)
                print(f"Sütun {index + 1} seçildi.")
                # İlgili sütun seçimiyle ilgili işlemleri gerçekleştirin
        else:
            if index in self.selected_columns:
                self.selected_columns.remove(index)
                print(f"Sütun {index + 1} seçimi kaldırıldı.")
                # İlgili sütun seçimiyle ilgili işlemleri gerçekleştirin

    def select_all_columns(self):
        for checkbox in self.column_checkboxes:
            checkbox.setChecked(True)

    def deselect_all_columns(self):
        for checkbox in self.column_checkboxes:
            checkbox.setChecked(False)

    def process_variation_kod(self):
        selected_columns = self.get_selected_columns()

        if len(selected_columns) < 2:
            QMessageBox.warning(self, "Uyarı", "Lütfen 'STOKKODU' sütununu ve bir belirleyici sütunu seçin.")
            return

        stok_kodu_column = None
        belirleyici_column = None

        for column in selected_columns:
            if column.lower() in ["stokkodu", "stok kodu", "stok_kodu"]:
                stok_kodu_column = column
            else:
                belirleyici_column = column

        if not stok_kodu_column:
            QMessageBox.warning(self, "Uyarı", "Lütfen 'STOKKODU' sütununu seçin.")
            return

        if not belirleyici_column:
            QMessageBox.warning(self, "Uyarı", "Lütfen bir belirleyici sütun seçin.")
            return

        variation_kod_list = []
        count_dict = {}

        for index, row in self.data_frame.iterrows():
            stok_kodu = row[stok_kodu_column]
            belirleyici = row[belirleyici_column]

            if pd.notna(stok_kodu) and pd.notna(belirleyici):
                data = f"{stok_kodu}{belirleyici}".replace(" ", "")
                count_dict[data] = count_dict.get(data, 0) + 1
                variation_kod_list.append(f"{data}-{count_dict[data]}")
            else:
                variation_kod_list.append("")  # Sonuç yoksa boş bir dize ekleyin

        self.variation_kod_results = variation_kod_list
        self.show_results()

    def process_breadcrumb_kat(self):
        selected_columns = self.get_selected_columns()

        if not selected_columns:
            QMessageBox.warning(self, "Uyarı", "Lütfen en az bir sütun seçin.")
            return

        breadcrumb_list = []

        for index, row in self.data_frame.iterrows():
            breadcrumb = ">".join(str(row[column]) for column in selected_columns if pd.notna(row[column]))
            breadcrumb_list.append(breadcrumb)

        self.breadcrumb_kat_results = breadcrumb_list
        self.show_results()

        # Bu işlem fonksiyonunun sonuna aşağıdaki kodu ekleyin:
        for i, breadcrumb in enumerate(self.breadcrumb_kat_results):
            self.data_frame.loc[i, 'Breadcrumb_Kat'] = breadcrumb

    def get_variations(self, selected_columns):
        if not self.csv_file_path:
            QMessageBox.warning(self, "Uyarı", "Önce bir CSV dosyası seçin.")
            return []

        try:
            # Read the CSV file using pandas
            self.data_frame = pd.read_csv(self.csv_file_path)

            variations = []
            for index, row in self.data_frame.iterrows():
                variation = {}
                for col_name in selected_columns:
                    cell_value = row[col_name]
                    variation[col_name] = cell_value
                if all(pd.notna(value) for value in variation.values()):  # Check if all values are non-empty
                    variations.append(variation)
            return variations

        except Exception as e:
            QMessageBox.critical(self, "Hata", "CSV dosyası okunurken bir hata oluştu:\n" + str(e))
            return []

    def process_variation(self):
        selected_columns = self.get_selected_columns()
        if not selected_columns:
            QMessageBox.warning(self, "Uyarı", "Lütfen bir veya daha fazla sütun seçin.")
            return

        variations = self.get_variations(selected_columns)

        variation_str_list = []
        variation_set = set()

        for variation in variations:
            combinations = list(itertools.product(*[variation[key].split(";") for key in selected_columns]))
            for combination in combinations:
                variation_items = [f"{selected_columns[i]};{combination[i]}" for i in range(len(selected_columns))]
                variation_str = ",".join(variation_items)

                if variation_str not in variation_set:
                    variation_set.add(variation_str)
                    variation_str_list.append(variation_str)

        self.variation_results = variation_str_list
        self.show_results()

        # Bu işlem fonksiyonunun sonuna aşağıdaki kodu ekleyin:
        for i, variation in enumerate(self.variation_results):
            self.data_frame.loc[i, 'Variation'] = variation

    def process_categories_1(self):
        selected_columns = self.get_selected_columns()

        if len(selected_columns) < 3:
            QMessageBox.warning(self, "Uyarı", "Lütfen en az 3 sütun seçin.")
            return

        kategori_sutun = selected_columns[0]
        alt_kategori_sutunlari = selected_columns[1:-1]
        urun_sutunu = selected_columns[-1]

        categories_str_list = []

        for index, row in self.data_frame.iterrows():
            kategori = row[kategori_sutun]
            alt_kategoriler = [row[sutun] for sutun in alt_kategori_sutunlari if pd.notna(row[sutun])]
            urun = row[urun_sutunu]

            if pd.notna(kategori) and alt_kategoriler and pd.notna(urun):
                breadcrumb = f"{kategori}>{'>'.join(alt_kategoriler)}; {urun}"
                categories_str_list.append(breadcrumb)

        self.categories_results = categories_str_list
        self.show_results()

        # Bu işlem fonksiyonunun sonuna aşağıdaki kodu ekleyin:
        for i, categories in enumerate(self.categories_results):
            self.data_frame.loc[i, 'Categories'] = categories

    def process_categories_2(self):
        selected_columns = self.get_selected_columns()

        if len(selected_columns) != 2:
            QMessageBox.warning(self, "Uyarı", "Lütfen tam olarak 2 sütun seçin.")
            return

        urun_sutunu = selected_columns[0]
        alt_kategori_sutunu = selected_columns[1]

        categories_str_list = []

        for index, row in self.data_frame.iterrows():
            urun = row[urun_sutunu]
            alt_kategori = row[alt_kategori_sutunu]

            if pd.notna(urun) and pd.notna(alt_kategori):
                categories_str_list.append(f"{urun}>{alt_kategori}")

        categories_result = ';'.join(categories_str_list)
        self.categories_results = [categories_result]
        self.show_results()

        # Bu işlem fonksiyonunun sonuna aşağıdaki kodu ekleyin:
        for i, categories in enumerate(self.categories_results):
            self.data_frame.loc[i, 'Categories'] = categories

    def show_results(self):
        # Clear previous result tabs
        self.operations_tabs.clear()

        if self.variation_kod_results:
            table_variation_kod = self.create_result_table(self.variation_kod_results)
            self.operations_tabs.addTab(table_variation_kod, "VARYASYONKODU")

        if self.breadcrumb_kat_results:
            table_breadcrumb_kat = self.create_result_table(self.breadcrumb_kat_results)
            self.operations_tabs.addTab(table_breadcrumb_kat, "BREADCRUMBKAT")

        if self.variation_results:
            table_variation = self.create_result_table(self.variation_results)
            self.operations_tabs.addTab(table_variation, "VARYASYON")

        if self.categories_results:
            table_categories = self.create_result_table(self.categories_results)
            self.operations_tabs.addTab(table_categories, "KATEGORİLER")


    def create_result_table(self, results):
        table = QTableWidget()
        table.setColumnCount(1)
        table.setRowCount(len(results))
        table.setHorizontalHeaderLabels(["Sonuçlar"])

        for row, result in enumerate(results):
            item = QTableWidgetItem(result)
            table.setItem(row, 0, item)

        table.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)  # Tüm sütunları genişlet

        return table
    
    def save_results_to_excel(self):
        if not self.csv_file_path:
            QMessageBox.warning(self, "Uyarı", "Önce bir CSV dosyası seçin.")
            return

        try:
            # Sonuç dizileri için bir sözlük oluşturun
            result_data = {
                "VARYASYONKODU": self.variation_kod_results,
                "BREADCRUMBKAT": self.breadcrumb_kat_results,
                "VARYASYON": self.variation_results,
                "KATEGORİLER": self.categories_results
            }

            # Sonuç dizileri arasındaki maksimum uzunluğu belirleyin
            max_length = max(len(arr) for arr in result_data.values())

            # Kısa dizileri maksimum uzunluğa ulaşmak için boş dizilerle doldurun
            for key, arr in result_data.items():
                if len(arr) < max_length:
                    result_data[key] += [""] * (max_length - len(arr))

            # Güncellenmiş sonuç dizileriyle DataFrame oluşturun
            result_df = pd.DataFrame(result_data)

            # Mevcut CSV verilerini okuyun
            excel_data = pd.read_csv(self.csv_file_path, encoding="utf-8")

            # Sonuç DataFrame'ini orijinal verilerle birleştirin
            merged_data = pd.concat([excel_data, result_df], axis=1)

            # Yeni Excel dosyasını kaydedin
            new_excel_file_path = self.csv_file_path.replace(".csv", "_with_results.xlsx")
            merged_data.to_excel(new_excel_file_path, index=False)

            QMessageBox.information(self, "Bilgi", "İşlem sonuçları Excel dosyasına kaydedildi: " + new_excel_file_path)
        except Exception as e:
            QMessageBox.critical(self, "Hata", "Excel dosyasına kaydedilirken bir hata oluştu:\n" + str(e))

    def cancel_operations(self):
        # Clear all result data and tabs
        self.variation_kod_results = []
        self.breadcrumb_kat_results = []
        self.variation_results = []
        self.categories_results = []
        self.show_results()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = ExcelProcessorApp()
    sys.exit(app.exec_())