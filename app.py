import streamlit as st
import easyocr
from PIL import Image
import openpyxl
import re
import numpy as np

# --------------------------------
# APP TITLE
# --------------------------------

st.title("EnergyBae Solar Load Calculator")

st.markdown(
    "Upload electricity bill and generate solar Excel automatically."
)

# --------------------------------
# FILE UPLOAD
# --------------------------------

uploaded_file = st.file_uploader(
    "Upload Electricity Bill",
    type=["jpg", "jpeg", "png"]
)

# --------------------------------
# OCR READER
# --------------------------------

reader = easyocr.Reader(['en'], gpu=False)

# --------------------------------
# PROCESS FILE
# --------------------------------

if uploaded_file:

    # Open Image
    image = Image.open(uploaded_file)

    # Show Image
    st.image(image, caption="Uploaded Bill")

    # Convert image
    image_np = np.array(image)

    st.write("Processing bill... please wait")

    # OCR Extraction
    results = reader.readtext(
        image_np,
        detail=0
    )

    # Convert list to text
    text = " ".join(results)

    # --------------------------------
    # SHOW OCR TEXT
    # --------------------------------

    st.subheader("Extracted Text")

    st.write(text)

    # --------------------------------
    # CONSUMER NUMBER
    # --------------------------------

    consumer_number = "Not Found"

    consumer_match = re.search(
        r'\d{10,15}',
        text
    )

    if consumer_match:

        consumer_number = consumer_match.group()

    # --------------------------------
    # UNITS EXTRACTION
    # --------------------------------

    units = "Not Found"

    unit_patterns = [
        r'(\d+)\s*kWh',
        r'(\d+)\s*Units',
        r'Units\s*[:\-]?\s*(\d+)',
        r'Consumption\s*[:\-]?\s*(\d+)',
        r'Consumed\s*[:\-]?\s*(\d+)'
    ]

    for pattern in unit_patterns:

        match = re.search(
            pattern,
            text,
            re.IGNORECASE
        )

        if match:

            value = int(match.group(1))

            # Ignore huge wrong numbers
            if value < 5000:

                units = value

                break

    # Backup logic
    if units == "Not Found":

        possible_numbers = re.findall(
            r'\b\d{2,4}\b',
            text
        )

        filtered = []

        for num in possible_numbers:

            value = int(num)

            # Typical electricity unit range
            if 10 <= value <= 2000:

                filtered.append(value)

        if filtered:

            units = filtered[0]

    # --------------------------------
    # LOAD EXTRACTION
    # --------------------------------

    load = "Not Found"

    load_match = re.search(
        r'(\d+\.?\d*)\s*KW',
        text,
        re.IGNORECASE
    )

    if load_match:

        load = load_match.group(1)

    # --------------------------------
    # BILL AMOUNT EXTRACTION
    # --------------------------------

    amount = "Not Found"

    amount_patterns = [
        r'Bill Amount\s*[:\-]?\s*(\d+)',
        r'Current Bill\s*[:\-]?\s*(\d+)',
        r'Amount\s*[:\-]?\s*(\d+)'
    ]

    for pattern in amount_patterns:

        match = re.search(
            pattern,
            text,
            re.IGNORECASE
        )

        if match:

            amount = match.group(1)

            break

    # Backup logic
    if amount == "Not Found":

        number_matches = re.findall(
            r'\d{3,5}',
            text
        )

        filtered_numbers = []

        for num in number_matches:

            value = int(num)

            if 100 <= value <= 10000:

                filtered_numbers.append(value)

        if filtered_numbers:

            amount = max(filtered_numbers)

    # --------------------------------
    # SHOW EXTRACTED DATA
    # --------------------------------

    st.subheader("Extracted Important Data")

    st.write("Consumer Number:", consumer_number)

    st.write("Units:", units)

    st.write("Load (KW):", load)

    st.write("Bill Amount:", amount)

    # --------------------------------
    # OPEN EXCEL TEMPLATE
    # --------------------------------

    workbook = openpyxl.load_workbook(
        "solar_template.xlsx"
    )

    sheet = workbook.active

    # --------------------------------
    # FILL EXCEL CELLS
    # --------------------------------

    sheet["B2"] = consumer_number
    sheet["B3"] = units
    sheet["B4"] = load
    sheet["B5"] = amount

    # --------------------------------
    # SAVE OUTPUT FILE
    # --------------------------------

    output_file = "filled_solar_output.xlsx"

    workbook.save(output_file)

    # --------------------------------
    # SUCCESS MESSAGE
    # --------------------------------

    st.success(
        "Excel Generated Successfully!"
    )

    # --------------------------------
    # DOWNLOAD BUTTON
    # --------------------------------

    with open(output_file, "rb") as file:

        st.download_button(
            label="Download Filled Excel",
            data=file,
            file_name=output_file
        )
