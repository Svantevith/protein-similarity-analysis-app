from io import StringIO
import gzip
import pandas as pd
import streamlit as st
from PIL import Image

from fasta_error import FASTAFormattingError
from helper_functions import (
    entry_range_to_extract,
    extract_fasta_from_data,
    find_similarities,
    select_sequences,
    select_validation_option,
    display_results,
)

# Open icon image
icon = Image.open("Images\dna_icon.ico")

# Set page configurations
st.set_page_config(
    page_title="Nucleotide Sequences Collinearity", page_icon=icon, layout="wide"
)

# Set page header image
header_img = Image.open("Images\istock_56013860_molecule_computer_2500.jpg")
st.image(header_img, use_column_width=True)
st.write(
    """
    # Nucleotide Sequences Collinearity
    Algorithm to identify the similarity between two FASTA queries.
    ***
    """
)

# ==============================
#    Accepted Input Formats
# ==============================
st.write(
    """
    ## FASTA
    A sequence in FASTA format begins with a single-line description, followed by lines of sequence data. 
    The description line (defline) is distinguished from the sequence data by a greater-than (">") symbol at the beginning. 
    
    > It is recommended that all lines of text be shorter than 80 characters in length.
      
    An **example sequence** in **FASTA format**:

        >P01013 GENE X PROTEIN (OVALBUMIN-RELATED)
        QIKDLLVSSSTDLDTTLVLVNAIYFKGMWKTAFNAEDTREMPFHVTKQESKPVQMMCMNNSFNVATLPAE
        KMKILELPFASGDLSMLVLLPDEVSDLERIEKTINFEKLTEWTNPNTMEKRRVKVYLPQMKIEEKYNLTS
        VLMALGMTDLFIPSANLTGISSAESLKISQAVHGAFMELSEDGIEMAGSTGVIEDIKHSPESEQFRADHP
        FLFLIKHNPTNTIVYFGRYWSP
    
    ***
    """
)

# ==============================
#    Validation types
# ==============================
validation_option, is_cross_val = select_validation_option()
excel_name = (
    "_".join(["FASTA", *map(str.title, validation_option.split(" "))]) + ".xlsx"
)

# =========================
#        FILE UPLOAD
# =========================
upload_checkbox = st.checkbox("I want to upload my own file")
if upload_checkbox:

    # Get entry ranges to extract
    min_idx, max_idx, indices, id_keys = entry_range_to_extract()

    upload_form = st.form(key="upload_input")
    upload_form.header("Upload file")
    upload_form.info(
        """
    File must be compliant with international FASTA strandards.
    
    For more information head to [National Center for Biotechnology Information](https://blast.ncbi.nlm.nih.gov/Blast.cgi?CMD=Web&PAGE_TYPE=BlastDocs&DOC_TYPE=BlastHelp)
    """
    )
    uploaded_file = upload_form.file_uploader(
        label="Only .txt, .zip and .gz files are supported", type=["txt", "zip", "gz"]
    )

    if uploaded_file:
        try:
            if uploaded_file.type == "text/plain":
                # To read file as bytes:
                bytes_data = uploaded_file.getvalue()

                # To convert to a string based IO:
                stringio = StringIO(bytes_data.decode("utf-8"))

                # Extract data from text file
                extracted_data = extract_fasta_from_data(
                    stringio, min_idx, max_idx, indices, id_keys
                )

            elif uploaded_file.type == "application/x-gzip":
                # Extract data from zipped file
                with gzip.open(uploaded_file, mode="rt", encoding="utf-8") as stringio:
                    extracted_data = extract_fasta_from_data(
                        stringio, min_idx, max_idx, indices, id_keys
                    )

            if not extracted_data.empty:
                # Select sequences by ID
                sequences = select_sequences(extracted_data)

                # Compute Summary Statistics and Similarity Matrix
                summary_statistics, similarity_matrix = find_similarities(
                    *sequences["SEQUENCE"], is_cross_val
                )

                # Display results section
                display_results(
                    excel_name,
                    sequences,
                    summary_statistics,
                    similarity_matrix,
                    validation_option,
                    is_cross_val,
                )

            else:
                st.warning("No FASTA data found")

        except FASTAFormattingError as fasta_error:
            st.error(str(fasta_error))
        except IndexError:
            st.error("Index not found in axis, out of range")
    else:
        st.warning("No file uploaded")

    upload_submit = upload_form.form_submit_button(label="Extract data")

else:
    # =========================
    #        CUSTOM INPUT
    # =========================
    error_msg = ""
    input_form = st.form(key="custom_input")
    input_form.header("Enter protein sequences")
    seq_1 = input_form.text_input(label="Write down first FASTA").strip()
    seq_2 = input_form.text_input(label="Write down second FASTA").strip()

    if seq_1 and seq_2:
        # Store Sequences in a DataFrame
        sequences = pd.DataFrame({"SEQUENCES": [seq_1, seq_2]})

        # Compute Summary Statistics and Similarity Matrix
        summary_statistics, similarity_matrix = find_similarities(
            seq_1, seq_2, is_cross_val
        )

        # Display results section
        display_results(
            excel_name,
            sequences,
            summary_statistics,
            similarity_matrix,
            validation_option,
            is_cross_val,
        )

    input_submit = input_form.form_submit_button(label="Approve sequences")
    if input_submit:
        if not seq_1:
            input_form.error("First FASTA cannot be blank")

        if not seq_2:
            input_form.error("Second FASTA cannot be blank")
