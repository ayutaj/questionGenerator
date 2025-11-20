import os
from helper import to_camel_case
import time

# ======================
# Question Paper Config
# ======================
# Topic heading shown at the top of the first page
TOPIC_NAME = f"ChangeName-{time.time()}"
# Input/Output file names
INPUT_FILE_NAME = "newQuestion.xlsx"
tab=4
tab2=4

#========================
#
# WORD DOCUMENT OPTIONS
#
#========================
#FORMAT OPTIONS
FORMAT=1  # 1: Two-column questions, answers at end; 2: Single-column questions with answers below each question

# Fonts
QUESTION_FONT = "Times New Roman"
QUESTION_FONT_SIZE = 12
OPTION_FONT = "Times New Roman"
OPTION_FONT_SIZE = 11
HEADING_FONT = "Arial Black"
HEADING_FONT_SIZE = 20

# Margins (in cm)
TOP_MARGIN_CM = 1
BOTTOM_MARGIN_CM = 1
LEFT_MARGIN_CM = 1
RIGHT_MARGIN_CM = 1

# Line spacing (1.0 = single, 1.15 = Word default)
LINE_SPACING = 1
tab=3

# ======================
# PowerPoint Formatting
# ======================
# PPT Fonts
PPT_QUESTION_FONT = "Times New Roman"     # Title (question) font
PPT_QUESTION_FONT_SIZE = 16

PPT_OPTION_FONT = "Times New Roman"        # Options font
PPT_OPTION_FONT_SIZE = 16



# ====================== DONT CHANGE BELOW THIS LINE ======================

INPUT_FOLDER = "input"
INPUT_FILE = os.path.join(INPUT_FOLDER, INPUT_FILE_NAME)

outputfoldername=to_camel_case(TOPIC_NAME)
OUTPUT_FOLDER = f"output/{outputfoldername}"
OUTPUT_QUESTIONS_FILE = os.path.join(OUTPUT_FOLDER, f"{outputfoldername}_questions.docx")
OUTPUT_ANSWERS_FILE = os.path.join(OUTPUT_FOLDER, f"{outputfoldername}_answer_key.docx")
OUTPUT_PPT_FILE = os.path.join(OUTPUT_FOLDER, f"{outputfoldername}_questions.pptx")


