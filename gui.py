from PySide6.QtWidgets import QMainWindow, QPushButton, QLabel, QLineEdit, QFileDialog, QVBoxLayout, QHBoxLayout, QWidget, QTextEdit,QColorDialog
from utils.read_excel import read_excel_range
from PyPDF2 import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics
from reportlab.lib.colors import HexColor
import io

class NameplateGeneratorGUI(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Nameplate Generator")

        # 파일 선택 관련 위젯들
        self.file_path_label = QLabel("참가자 명단 (xls, xlsx)")
        self.file_path_input = QLineEdit()
        self.file_select_button = QPushButton("Select File")
        self.file_select_button.clicked.connect(self.select_file)

        self.background_label = QLabel("배경 이미지 (pdf)")
        self.background_input = QLineEdit()
        self.background_select_button = QPushButton("Select Image")
        self.background_select_button.clicked.connect(self.select_background_image)

        self.font_path_label = QLabel("폰트 (ttf)")
        self.font_path_input = QLineEdit()
        self.font_path_button = QPushButton("Select Font")
        self.font_path_button.clicked.connect(self.select_font_path)

        self.hex_color_label = QLabel("색상코드 (HEX)")
        self.hex_color_input = QLineEdit()

        # 조 이름, 범위, 결과물 저장 경로, 생성 버튼, 로그 위젯
        self.team_name_label = QLabel("조이름")
        self.team_name_input = QLineEdit()

        self.range_label = QLabel("조원 이름 및 학번 (B2:C14)")
        self.range_input = QLineEdit()

        self.save_path_label = QLabel("결과물 저장 경로")
        self.save_path_input = QLineEdit()
        self.save_path_button = QPushButton("Select Folder")
        self.save_path_button.clicked.connect(self.select_save_folder)

        self.generate_button = QPushButton("명찰 생성")
        self.generate_button.clicked.connect(self.generate_nameplate)

        self.log_label = QLabel("Log:")
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)

        # 조 이름, 학번, 이름 위치 및 폰트 크기 입력 위젯들
        self.team_name_position_label = QLabel("조이름 세로 위치(mm)")
        self.team_name_position_input = QLineEdit()

        self.team_name_font_size_label = QLabel("조이름 폰트 크기(pt)")
        self.team_name_font_size_input = QLineEdit()

        self.student_id_position_label = QLabel("학번 세로 위치(mm)")
        self.student_id_position_input = QLineEdit()

        self.student_id_font_size_label = QLabel("학번 폰트 크기(pt)")
        self.student_id_font_size_input = QLineEdit()

        self.name_position_label = QLabel("이름 세로 위치(mm)")
        self.name_position_input = QLineEdit()

        self.name_font_size_label = QLabel("이름 폰트 크기(pt)")
        self.name_font_size_input = QLineEdit()

        # 레이아웃 구성
        file_layout = QHBoxLayout()
        file_layout.addWidget(self.file_path_label)
        file_layout.addWidget(self.file_path_input)
        file_layout.addWidget(self.file_select_button)

        background_layout = QHBoxLayout()
        background_layout.addWidget(self.background_label)
        background_layout.addWidget(self.background_input)
        background_layout.addWidget(self.background_select_button)

        font_path_layout = QHBoxLayout()
        font_path_layout.addWidget(self.font_path_label)
        font_path_layout.addWidget(self.font_path_input)
        font_path_layout.addWidget(self.font_path_button)

        color_layout = QHBoxLayout()
        color_layout.addWidget(self.hex_color_label)
        color_layout.addWidget(self.hex_color_input )

        team_layout = QHBoxLayout()
        team_layout.addWidget(self.team_name_label)
        team_layout.addWidget(self.team_name_input)

        range_layout = QHBoxLayout()
        range_layout.addWidget(self.range_label)
        range_layout.addWidget(self.range_input)

        team_name_layout = QHBoxLayout()
        team_name_layout.addWidget(self.team_name_position_label)
        team_name_layout.addWidget(self.team_name_position_input)
        team_name_layout.addWidget(self.team_name_font_size_label)
        team_name_layout.addWidget(self.team_name_font_size_input)

        student_id_layout = QHBoxLayout()
        student_id_layout.addWidget(self.student_id_position_label)
        student_id_layout.addWidget(self.student_id_position_input)
        student_id_layout.addWidget(self.student_id_font_size_label)
        student_id_layout.addWidget(self.student_id_font_size_input)

        name_layout = QHBoxLayout()
        name_layout.addWidget(self.name_position_label)
        name_layout.addWidget(self.name_position_input)
        name_layout.addWidget(self.name_font_size_label)
        name_layout.addWidget(self.name_font_size_input)

        save_path_layout = QHBoxLayout()
        save_path_layout.addWidget(self.save_path_label)
        save_path_layout.addWidget(self.save_path_input)
        save_path_layout.addWidget(self.save_path_button)

        # 전체 메인 레이아웃
        main_layout = QVBoxLayout()
        main_layout.addLayout(file_layout)
        main_layout.addLayout(background_layout)
        main_layout.addLayout(font_path_layout)
        main_layout.addLayout(color_layout)
        main_layout.addLayout(team_layout)
        main_layout.addLayout(range_layout)
        main_layout.addLayout(team_name_layout)
        main_layout.addLayout(student_id_layout)
        main_layout.addLayout(name_layout)
        main_layout.addLayout(save_path_layout)
        main_layout.addWidget(self.generate_button)
        main_layout.addWidget(self.log_label)
        main_layout.addWidget(self.log_text)

        central_widget = QWidget()
        central_widget.setLayout(main_layout)
        self.setCentralWidget(central_widget)

    def select_file(self):
        options = QFileDialog.Options()
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Select Excel File", "", "Excel Files (*.xlsx *.xls)", options=options
        )
        if file_path:
            self.file_path_input.setText(file_path)

    def select_background_image(self):
        options = QFileDialog.Options()
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Select Background PDF", "", "PDF Files (*.pdf)", options=options
        )
        if file_path:
            self.background_input.setText(file_path)

    def select_font_path(self):
        options = QFileDialog.Options()
        font_path, _ = QFileDialog.getOpenFileName(
            self, "Select Font File", "", "Font Files (*.ttf)", options=options
        )
        if font_path:
            self.font_path_input.setText(font_path)
            pdfmetrics.registerFont(TTFont("CustomFont", font_path))

    def select_save_folder(self):
        options = QFileDialog.Options()
        folder_path = QFileDialog.getExistingDirectory(self, "Select Folder")
        if folder_path:
            self.save_path_input.setText(folder_path)
    def select_text_color(self):
        color_dialog = QColorDialog()
        color = color_dialog.getColor()
        if color.isValid():
            self.text_color = color.name()
            # 모든 텍스트 필드의 색상 변경
            self.set_text_field_color(self.team_name_input)
            self.set_text_field_color(self.student_id_input)
            self.set_text_field_color(self.name_input)

    def set_text_field_color(self, text_field):
        text_field.setStyleSheet(f"color: {self.text_color}")

    def generate_nameplate(self):
        # 사용자 입력 데이터 정리
        file_path = self.file_path_input.text()
        background_path = self.background_input.text()
        font_path = self.font_path_input.text()
        team_name = self.team_name_input.text()
        cell_range = self.range_input.text()
        save_path = self.save_path_input.text()
        color = self.hex_color_input.text()
        try:
            team_name_position = float(self.team_name_position_input.text())
            team_name_font_size = int(self.team_name_font_size_input.text())
            student_id_position = float(self.student_id_position_input.text())
            student_id_font_size = int(self.student_id_font_size_input.text())
            name_position = float(self.name_position_input.text())
            name_font_size = int(self.name_font_size_input.text())

        except ValueError:
            self.log_text.append("텍스트 위치와 크기 필드에는 올바른 값를 입력하세요.")
            return

        if not all([file_path, background_path, font_path, team_name, cell_range, save_path,
                    team_name_position, team_name_font_size, student_id_position, student_id_font_size,
                    name_position, name_font_size]):
            self.log_text.append("모든 필드를 채워주세요.")
            return
        
        try:
            result = read_excel_range(file_path, cell_range)
            name_data, id_data = result
            self.log_text.append("이름 데이터:")
            self.log_text.append(", ".join(str(cell) for cell in name_data))

            self.log_text.append("학번 데이터:")
            self.log_text.append(", ".join(str(cell) for cell in id_data))  # 데이터를 로그에 출력

            background_pdf = PdfReader(background_path)

            page_width = float(background_pdf.pages[0].mediabox[2])
            page_height = float(background_pdf.pages[0].mediabox[3])

            # PDF에 데이터 합성
            output = PdfWriter()

            for name, student_id in zip(name_data, id_data):
                packet = io.BytesIO()
                can = canvas.Canvas(packet, pagesize=letter)

                # 여기서 조이름, 학번, 이름을 각각의 위치와 폰트 크기로 조정하여 합성하세요
                team_text = team_name # 여기에 실제 조이름 변수 입력
                student_id_text = str(student_id)  # 여기에 실제 학번 변수 입력
                name_text = name  # 여기에 실제 이름 변수 입력

                # 가운데 정렬을 위한 x 좌표 설정

                center_x = page_width / 2
                can.setFillColor(HexColor(color))
                # 조이름 합성
                can.setFont("CustomFont", team_name_font_size)
                can.drawString(center_x - can.stringWidth(team_text) / 2,  page_height - (team_name_position * 2.9), team_text)

                # 학번 합성
                can.setFont("CustomFont", student_id_font_size)
                can.drawString(center_x - can.stringWidth(student_id_text) / 2,  page_height - (student_id_position * 2.9), student_id_text)

                # 이름 합성
                can.setFont("CustomFont", name_font_size)
                can.drawString(center_x - can.stringWidth(name_text) / 2,  page_height - (name_position * 2.9), name_text)

                can.save()

                # Move to the beginning of the StringIO buffer
                packet.seek(0)
                new_pdf = PdfReader(packet)
                existing_pdf = PdfReader(background_path)
                output = PdfWriter()

                # Merge the PDFs
                background_pdf = existing_pdf.pages[0]
                background_pdf.merge_page(new_pdf.pages[0])
                output.add_page(background_pdf)

                # Save the output PDF to the desired location
                output_path = f"{save_path}/{name_text}.pdf"
                with open(output_path, "wb") as out_file:
                    output.write(out_file)

            self.log_text.append("명찰 생성 완료: " + output_path)

        except Exception as e:
            self.log_text.append(f"오류 발생: {e}")