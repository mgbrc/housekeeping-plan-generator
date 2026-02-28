import os
import re
import glob
import pandas as pd
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.platypus import (
    SimpleDocTemplate,
    Table,
    TableStyle,
    Spacer,
    KeepInFrame,
    Paragraph,
    Flowable,
)
from reportlab.lib.styles import getSampleStyleSheet
from datetime import datetime

# Custom Flowable for drawing adaptive right-side cell content
class RightCellAdaptive(Flowable):

    def __init__(self, bed_letter, status):
        super().__init__()
        self.bed_letter = bed_letter
        self.status = status

    def wrap(self, availWidth, availHeight):
        self.W = availWidth
        self.H = availHeight
        return availWidth, availHeight

    def draw(self):
        W, H = self.W, self.H

        # Draw bed letter (scaled to fit row height)
        if self.bed_letter:
            size = H * 0.8
            self.canv.setFont("Helvetica", size)
            y = H / 2 - size * 0.35
            self.canv.drawCentredString(W * 0.25, y, self.bed_letter)

        # Stayover → draw "+"
        if self.status == "Stayover":
            size = H * 0.96
            self.canv.setFont("Helvetica-Bold", size)
            y = H / 2 - size * 0.35
            self.canv.drawCentredString(W * 0.6, y, "+")

        # Turnover / Check-out → draw filled block
        elif self.status in ("Turnover", "Check-out"):
            block_w = W * 0.35
            block_h = H * 0.88
            x = W - block_w
            y = (H - block_h) / 2
            self.canv.rect(x, y, block_w, block_h, fill=1, stroke=0)

# Output PDF will be saved to Desktop
desktop = os.path.join(os.path.expanduser("~"), "Desktop")
output_file = os.path.join(desktop, "plan_symbol_version.pdf")

# Automatically detect latest Excel file in Downloads
downloads = os.path.join(os.path.expanduser("~"), "Downloads")
excel_files = glob.glob(os.path.join(downloads, "*.xlsx"))

if not excel_files:
    raise FileNotFoundError(
        "No Excel (.xlsx) files found in your Downloads folder."
    )

# Select most recently modified Excel file
latest_excel = max(excel_files, key=os.path.getmtime)
print(f"Using the latest downloaded Excel file: {latest_excel}")
excel_file = latest_excel

# Room categories (used to determine number of beds)
ten_bed_rooms = ["10", "20", "30", "40", "50"]
six_bed_rooms = ["16", "26", "36", "45", "55"]
four_bed_rooms = ["17", "18", "27", "28", "37", "38", "41", "42", "51", "52"]
three_bed_rooms = ["19", "29", "39", "43", "44", "53", "54"]
private_double_rooms = ["12", "14", "22", "24", "32", "34"]
private_twin_rooms = ["11", "13", "15", "21", "23", "25", "31", "33", "35"]

# Load and preprocess Excel data
df = pd.read_excel(excel_file, header=0)
# Keep only ROOM and FRONTDESK STATUS columns (by index)
df = df.iloc[:, [1, 5]]
df.columns = ["ROOM", "FRONTDESK STATUS"]

# Dictionaries to store statuses
room_bed_status = {}
room_room_status = {}

for _, row in df.iterrows():
    raw_room = str(row["ROOM"]).strip()

    # Skip empty rows
    if raw_room in ("nan", ""):
        continue

    status = (
        str(row["FRONTDESK STATUS"]).strip()
        if pd.notna(row["FRONTDESK STATUS"])
        else ""
    )

    # Ignore irrelevant status
    if status in ("Not Reserved"):
        continue

    parts = raw_room.split("-")
    building = parts[0] if len(parts) > 1 else ""
    has_star = "(*)" in raw_room

    # Extract two-digit room number
    m_room = re.search(r"(\d{2})", raw_room)
    if not m_room:
        continue

    room_num = m_room.group(1)

    m_bed = re.search(r"-(?:\(\*\))?-?([A-J])$", raw_room)
    bed_letter = m_bed.group(1) if m_bed else None

    # If specific bed
    if bed_letter:
        key = f"{room_num}-{bed_letter}"
        room_bed_status[key] = status
        continue

    # If entire room (marked with (*))
    if has_star:
        room_room_status[room_num] = status
        continue

# PDF generation
doc = SimpleDocTemplate(output_file, pagesize=landscape(A4))
elements = []

# Room sections (each column block)
sections = [
    list(range(10, 19)),
    list(range(20, 29)),
    list(range(30, 39)),
    list(range(40, 45)),
    list(range(50, 55)),
]

col_widths = [25, 35]
row_height = 12
all_private_rooms = private_double_rooms + private_twin_rooms
section_frames = []

# Build each section (column)
for sec in sections:
    col_tables = []

    for room_num in sec:
        room_str = str(room_num)

        # Determine number of beds
        if room_str in all_private_rooms:
            beds = []
        elif room_str in ten_bed_rooms:
            beds = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J"]
        elif room_str in six_bed_rooms:
            beds = ["A", "B", "C", "D", "E", "F"]
        elif room_str in four_bed_rooms:
            beds = ["A", "B", "C", "D"]
        elif room_str in three_bed_rooms:
            beds = ["A", "B", "C"]
        else:
            continue

        data = []

        # Dorm rooms (multiple beds)
        if beds:
            for idx, bed in enumerate(beds):
                key = f"{room_str}-{bed}"
                status = (
                    room_bed_status.get(key)
                    or room_room_status.get(room_str)
                )
                right_cell = RightCellAdaptive(bed, status)
                left_cell = room_str if idx == 0 else ""
                data.append([left_cell, right_cell])

        # Private rooms
        else:
            status = room_room_status.get(room_str)

            if status in ("Turnover", "Check-out") or status == "Stayover":
                right_cell = RightCellAdaptive("", status)
            else:
                right_cell = ""

            data.append([room_str, right_cell])

        # Create table for room
        t = Table(
            data,
            colWidths=col_widths,
            rowHeights=[row_height] * len(data),
        )

        # Styling
        t.setStyle(
            TableStyle(
                [
                    ("GRID", (0, 0), (-1, -1), 0.5, colors.black),
                    ("ALIGN", (0, 0), (0, -1), "CENTER"),
                    ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                    ("LEFTPADDING", (1, 0), (1, -1), 0),
                    ("RIGHTPADDING", (1, 0), (1, -1), 0),
                    ("TOPPADDING", (1, 0), (1, -1), 0),
                    ("BOTTOMPADDING", (1, 0), (1, -1), 0),
                    ("BACKGROUND", (0, 0), (0, 0), colors.whitesmoke),
                    ("FONTSIZE", (0, 0), (-1, -1), 8),
                ]
            )
        )

        col_tables.append(t)
        col_tables.append(Spacer(1, 2))

    frame = KeepInFrame(
        maxWidth=800,
        maxHeight=A4[1] - 10,
        content=col_tables,
        hAlign="LEFT",
        vAlign="TOP",
    )

    section_frames.append(frame)

# Footer with timestamp
styles = getSampleStyleSheet()
footer_style = styles["Normal"]
footer_style.fontSize = 12
footer_style.alignment = 0

current_time = datetime.now().strftime("%d.%m.%Y, %H:%M")
footer_text = Paragraph(f"Generated on: {current_time}", footer_style)

elements.append(footer_text)

# Main layout table (5 columns)
main_table = Table(
    [section_frames],
    colWidths=[90] * 5,
    hAlign="LEFT",
)

main_table.setStyle(
    TableStyle(
        [
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ]
    )
)

elements.append(main_table)

# Build PDF
doc.build(elements)

print(f"PDF plan successfully created: {output_file}")
input("Enter")