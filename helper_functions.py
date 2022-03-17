import re
import pythoncom
import numpy as np
import pandas as pd
import seaborn as sns
import streamlit as st
from pathlib import Path
import altair as alt
from operator import itemgetter
import win32com.client as win32
from io import TextIOWrapper, BytesIO
from typing import Sequence, Union, Tuple
from xlsxwriter.worksheet import Worksheet
from streamlit.delta_generator import DeltaGenerator
import matplotlib.pyplot as plt
from PIL import Image

from fasta_error import FASTAFormattingError

# Column naming for Sequences DataFrame
FASTA_COLUMNS = ["ID", "DESCRIPTION", "SEQUENCE"]


def select_validation_option() -> Tuple[str, bool]:
    st.write(
        """
    ## Similarity Matrix is validated for collinearity
    """
    )

    validation_option = st.radio(
        label="How to compare the sequences?",
        options=(
            "Diagonal validation",
            "Cross validation",
        ),
    )

    cols = st.columns(2)
    for i, img_path in enumerate(
        ["Images/diagonal_validation.png", "Images/cross_validation.png"]
    ):
        validation_img = Image.open(img_path)
        cols[i].image(validation_img, use_column_width=True)

    if validation_option == "Diagonal validation":
        is_cross_val = False

    is_cross_val = validation_option == "Cross validation"
    st.write("***")

    return validation_option, is_cross_val


# Display all results including the histogram, table and download option
def display_results(
    excel_name: str,
    sequences: pd.DataFrame,
    summary_statistics: pd.DataFrame,
    similarity_matrix: pd.DataFrame,
    title: str,
    is_cross_val: bool,
) -> None:
    st.write(
        f"""
        ***
        ## {title}
        Investigate distribution and position matrix of similarities between the input sequences.
        ***
        """
    )

    cols = st.columns(2)

    # Adjust Similarity Matrix with a slider
    cols[0].write(
        """
        ### Discover collinear nucleotides
        Examine **particular indices** at which the **nodes overlap**.
        """
    )
    slider_table_display(similarity_matrix, is_cross_val, column=cols[0])

    # Graphically present Summary Statistics on a bar plot
    cols[1].write(
        """
        ### Analyze histogram
        Get deeper insights on the **distribution of similarities** for **each repeating nucleotide**.
        ##
        """
    )
    plot_summary_statistics(
        summary_statistics,
        title="Variation in collinearity by nucleotide",
        height=420,
        column=cols[1],
    )

    # Download Excel outcomes
    st.write(
        f"""
        ***
        ### Download spreadsheet
        **Excel** file contains all **information extracted from provided sequences**, evaluated **summary statistics**, and formatted **similarity matrix**.
        """
    )
    download_excel(
        excel_name,
        sequences,
        summary_statistics,
        similarity_matrix,
        title,
        is_cross_val,
    )


# Plot histogram for similarities
def plot_summary_statistics(
    summary_statistics: pd.DataFrame,
    height: int = -1,
    title: str = "",
    column: DeltaGenerator = None,
) -> None:
    # Prepare properties params
    properties = {
        "title": title,
        "autosize": alt.AutoSizeParams(type="fit", contains="padding"),
    }
    if height > 0:
        properties["height"] = height

    bar_chart = (
        alt.Chart(summary_statistics)
        .mark_bar()
        .encode(
            alt.X("Nucleotide", type="ordinal", axis=alt.Axis(labelAngle=0)),
            alt.Y("Similarities", type="quantitative"),
            alt.Color(
                "Nucleotide", type="ordinal", scale=alt.Scale(scheme="redyellowblue")
            ),
            alt.Tooltip(["Nucleotide", "Similarities"]),
        )
        .properties(**properties)
        .interactive()
    )
    if column:
        column.altair_chart(bar_chart, use_container_width=True)
    else:
        st.altair_chart(bar_chart, use_container_width=True)


# Present similarities on the scatter plot
def plot_similarities(similarity_matrix: pd.DataFrame, title: str) -> None:
    # Create Figure
    fig = plt.figure()

    # Scatter plot
    ## Find coordinates of matching fields
    x, y = np.where(similarity_matrix == 1)

    ## Prepare colors
    color_keys = max(
        [np.unique(similarity_matrix.index), np.unique(similarity_matrix.columns)],
        key=len,
    )
    rgb_values = sns.color_palette("rocket", len(color_keys))
    colors = dict(zip(color_keys, rgb_values))

    ## Plot each matching nucleotide
    for i in range(x.shape[0]):
        c_idx = similarity_matrix.index[x[i]]
        label = (
            c_idx
            if c_idx not in plt.gca().get_legend_handles_labels()[1]
            else "_nolegend_"
        )
        plt.scatter(x[i], y[i], label=label, color=colors[c_idx])

    # Linear regression
    ## Define axes
    x, y = (np.arange(n) for n in similarity_matrix.shape)
    m, b = np.polyfit(x, y, 1)
    plt.plot(x, m * x + b, "k", zorder=0)

    plt.xticks(np.arange(similarity_matrix.shape[0]), similarity_matrix.index)
    plt.yticks(np.arange(similarity_matrix.shape[1]), similarity_matrix.columns)
    plt.title(title)
    plt.legend()

    # Display graph
    st.pyplot(fig)


# Select Sequences by ID
def select_sequences(extracted_data: pd.DataFrame) -> pd.DataFrame:
    st.write("### Select sequences by ID")
    selected = []
    labels = ["first", "second"]
    for i, col in enumerate(st.columns(2)):
        id = col.selectbox(f"Select {labels[i]} sequence", extracted_data["ID"], key=i)
        option = extracted_data.loc[extracted_data["ID"] == id]
        sequence = option["SEQUENCE"].iloc[0]
        selected.append(option)
        with col.expander("View sequence"):
            st.text(sequence)

    return pd.concat(selected, ignore_index=True)


# Slider control
def slider(limit: int, column: DeltaGenerator = None) -> Tuple[int, int]:
    slider_range = (1, limit)
    if column:
        min_res, max_res = column.slider(
            label="Change the similarity matrix display range",
            min_value=slider_range[0],
            max_value=slider_range[1],
            value=slider_range,
        )
    else:
        min_res, max_res = st.slider(
            label="Change the similarity matrix display range",
            min_value=slider_range[0],
            max_value=slider_range[1],
            value=slider_range,
        )

    min_res -= 1
    max_res += 1

    return min_res, max_res


# Metrics component
def display_metrics(
    similarity_matrix: pd.DataFrame,
    cross_val: bool = False,
    column: DeltaGenerator = None,
) -> None:
    n_similarities = (similarity_matrix.values == 1).sum()

    n_cols = similarity_matrix.shape[1]
    n_entries = n_cols ** 2 if cross_val else n_cols

    similarity_ratio = np.round(n_similarities * 100 / n_entries, 2)
    if column:
        column.metric(
            label="Similarities",
            value=f"{n_similarities} nucleic acids correspond",
            delta=f"{similarity_ratio}%",
            delta_color="off",
        )
    else:
        st.metric(
            label="Similarities",
            value=f"{n_similarities} nucleic acids correspond",
            delta=f"{similarity_ratio}%",
            delta_color="off",
        )


# Minipulate Similarity Matrix' dimensions using a slider control
def slider_table_display(
    similarity_matrix: pd.DataFrame,
    cross_val: bool = False,
    column: DeltaGenerator = None,
) -> None:
    slider_min, slider_max = slider(similarity_matrix.shape[1], column)
    similarity_results = similarity_matrix.iloc[
        slider_min:slider_max, slider_min:slider_max
    ]

    display_avoid_nan(similarity_results, column)
    display_metrics(similarity_results, cross_val, column)


# Return enumerated labels for DataFrame display
def enumerate_sequence_labels(labels: Union[str, list]) -> list:
    return [f"{i}-{c}" if c else f"{i}-N/A" for i, c in enumerate(labels, start=1)]


# Display Similarity Matrix with a proper formatting
def display_avoid_nan(df: pd.DataFrame, column=None) -> None:
    # Encode numerical & missing values
    encoded = pd.DataFrame(
        data=df.values,
        columns=enumerate_sequence_labels(df.columns),
        index=enumerate_sequence_labels(df.index),
    ).replace([1, 0, np.nan], ["🟢", "", ""])

    # Display encoded table
    column.write(encoded) if column else st.write(encoded)


# Validate naming convention for .xlsx files
def validate_xlsx(filename: str) -> str:
    return filename if filename.lower().endswith(".xlsx") else f"{filename}.xlsx"


# Return file's download path
def download_path(filename: str) -> Path:
    return Path(Path.home(), "Downloads", filename)


# Download Excel using a button
def download_excel(
    filename: str,
    sequences: pd.DataFrame,
    summary_statistics: pd.DataFrame,
    similarity_matrix: pd.DataFrame,
    title: str = "",
    cross_val: bool = False,
) -> bool:
    xlsx_name = validate_xlsx(filename)
    xlsx_path = download_path(xlsx_name)
    xlsx_data = process_excel(
        sequences, summary_statistics, similarity_matrix, title, cross_val
    )

    is_downloaded = st.download_button(
        label="Download",
        data=xlsx_data,
        file_name=xlsx_name,
        on_click=autofit_column_width,
        args=(xlsx_path,),
    )
    if is_downloaded:
        st.info(f"Excel successfully downloaded as {xlsx_path}")


# Automatically adjust Excel columns width
def autofit_column_width(path: Path):
    path.unlink(missing_ok=True)
    excel = win32.gencache.EnsureDispatch("Excel.Application", pythoncom.CoInitialize())
    excel.Visible = False
    wb = excel.Workbooks.Open(path)
    for ws in wb.Worksheets:
        ws.Columns.AutoFit()
    wb.Save()
    excel.Application.Quit()


# Apply XlsxWriter formatting to specified cells
def apply_formatting(
    ws: Worksheet, fmt: str, cells_range: Union[str, Sequence[int]]
) -> None:
    # Prepare parameter dictionary
    fmt_params = {"type": "no_errors", "format": fmt}

    # Row/Column notation
    if isinstance(cells_range, str):
        ws.conditional_format(
            cells_range,
            fmt_params,
        )

    # Cell notation
    else:
        ws.conditional_format(
            *cells_range,
            fmt_params,
        )


# Process Excel file so it is available for download
# IMPORTANT: Cache the conversion to prevent computation on every rerun
@st.cache  # (suppress_st_warning=True)
def process_excel(
    sequences: pd.DataFrame,
    summary_statistics: pd.DataFrame,
    similarity_matrix: pd.DataFrame,
    title: str = "",
    cross_val: bool = False,
) -> bytes:
    # Create temporary IO
    output = BytesIO()

    # Initiate ExcelWriter object
    writer = pd.ExcelWriter(output, engine="xlsxwriter")

    # Get the active workbook
    workbook = writer.book

    # Define primary & secondary formatting
    primary_fmt = workbook.add_format({"border": 1, "bold": True})
    secondary_fmt = workbook.add_format({"border": 1})

    # =====================================
    #               SHEET 1
    # =====================================

    # Sheet 1 - Sequences description
    sequences.to_excel(writer, index=False, sheet_name="Sequences")
    sheet_1 = writer.sheets["Sequences"]

    ## Apply formatting
    n_rows, n_cols = sequences.shape
    cells_range = (1, 0, n_rows, n_cols - 1)
    apply_formatting(sheet_1, secondary_fmt, cells_range)

    # =====================================
    #               SHEET 2
    # =====================================

    # Sheet 2 - Summary Statistics
    ## Export Summary Statistics
    summary_statistics.to_excel(writer, index=False, sheet_name="Summary Statistics")
    sheet_2 = writer.sheets["Summary Statistics"]

    ## Write down the Total number of similarities
    n_rows = summary_statistics.shape[0] + 1
    total = f"Total: {summary_statistics['Similarities'].sum()}"
    cells_range = f"A2:C{n_rows}"
    sheet_2.write(n_rows, 2, total, primary_fmt)
    apply_formatting(sheet_2, secondary_fmt, cells_range)

    # Create a chart object.
    chart = workbook.add_chart({"type": "column"})
    formula = "='Summary Statistics'!${}$2:${}$" + str(n_rows)

    # Get color palette
    cp = sns.color_palette("hls", n_rows - 1).as_hex()

    # Configure the series of the chart from the dataframe data.
    chart.add_series(
        {
            "name": "Similarities per nucleotide",
            "categories": formula.format("A", "A"),
            "values": formula.format("C", "C"),
            "gap": 50,
            "trendline": {
                "type": "moving_average",
                "name": "Moving average",
                "period": 2,
                "line": {
                    "color": "red",
                    "width": 1.25,
                },
            },
            "points": [{"fill": {"color": c}} for c in cp],
        }
    )

    # Configure the chart axes.
    chart.set_x_axis({"name": "Nucleotide"})
    chart.set_y_axis({"name": "Similarities"})
    chart.set_title({"name": title})

    # Turn off chart legend. It is on by default in Excel.
    chart.set_legend({"position": "none"})

    # Insert chart
    sheet_2.insert_chart(f"E2", chart)

    # =====================================
    #               SHEET 3
    # =====================================

    # Sheet 3 - Similarity Matrix
    similarity_matrix.to_excel(writer, sheet_name="Similarity Matrix")
    sheet_3 = writer.sheets["Similarity Matrix"]

    ## Define background formatting
    positive_fmt = workbook.add_format({"center_across": True, "bg_color": "green"})
    negative_fmt = workbook.add_format({"center_across": True, "bg_color": "red"})

    ## Apply formatting to the entire worksheet
    cells_range = (1, 1, *similarity_matrix.shape)
    apply_formatting(sheet_3, primary_fmt, cells_range)

    # Apply cell formatting for each cell
    if cross_val:
        for i in range(similarity_matrix.shape[0]):
            for j in range(similarity_matrix.shape[1]):
                val = similarity_matrix.iloc[i, j]
                cell_fmt = positive_fmt if val else negative_fmt
                cell_fmt.set_center_across()
                sheet_3.write(i + 1, j + 1, val, cell_fmt)

    ## Apply cell formatting along the diagonal
    else:
        for i in range(similarity_matrix.shape[1]):
            j = i + 1
            val = similarity_matrix.iloc[i, i]
            cell_fmt = positive_fmt if val else negative_fmt
            cell_fmt.set_center_across()
            sheet_3.write(j, j, val, cell_fmt)

    # Save & close ExcelWriter
    writer.save()

    # Process Excel data
    processed_data = output.getvalue()
    return processed_data


# Get Summary Statistics and Similarity Matrix
def find_similarities(
    seq_1: str, seq_2: str, cross_val: bool = False, stats: bool = False
) -> pd.DataFrame:
    # Return empty dataframe if any of sequences is empty
    if not (seq_1 and seq_2):
        return

    # Find shortest (y-axis) and longest (x-axis) sequence
    xx = max([seq_1, seq_2], key=len)
    yy = seq_2 if xx == seq_1 else seq_1

    # Store lengths as reusable variables
    n = len(xx)
    m = len(yy)

    # Convert sequences from string to list
    xx = list(xx)
    yy = list(yy)

    # Pad shorter sequence to equalize matrix dimensions
    if m < n:
        yy = np.pad(
            list(yy),
            pad_width=(0, n - m),
            mode="constant",
            constant_values="",
        )

    # Create initially empty matrix
    similarities = np.full(shape=(n, n), fill_value=np.nan)

    # Create initially empty dictionary
    statistics = {}
    for i, x in enumerate(xx):
        covers = []
        yy_enum = enumerate(yy) if cross_val else enumerate([yy[i]], start=i)
        for j, y in yy_enum:
            is_cover = x == y
            similarities[i, j] = int(is_cover)
            if is_cover:
                covers.append((i, j))

        if covers:
            try:
                statistics[x][0].extend(covers)
            except KeyError:
                statistics[x] = [covers]

    summary_statistics = pd.DataFrame.from_dict(
        statistics, orient="index", columns=["Position"]
    )
    summary_statistics["Similarities"] = summary_statistics["Position"].map(len)
    summary_statistics.reset_index(inplace=True)
    summary_statistics.rename({"index": "Nucleotide"}, axis=1, inplace=True)
    summary_statistics.sort_values(by=["Similarities"], ascending=False, inplace=True)

    similarity_matrix = pd.DataFrame(
        data=similarities,
        columns=xx,
        index=yy,
    )

    return summary_statistics, similarity_matrix


# Prompt user to reduce extracted data
def entry_range_to_extract() -> Tuple[int, int, list, list]:
    with st.expander("Expand to enter specific upload criteria"):
        # Inline double grid
        row_1 = st.columns(2)

        help_text = "{} can be separated by any delimeter"
        # Indices of entries to extract
        indices = [
            int(i)
            for i in re.findall(
                "\\d+", row_1[0].text_input("Entries", help=help_text.format("indices"))
            )
        ]

        # ID Keys of entries to extract
        id_keys = re.findall(
            "\\w+", row_1[1].text_input("ID Keys", help=help_text.format("IDs"))
        )

        # Inline triple grid
        row_2 = st.columns(2)

        # Index of first FASTA to extract (starts from 1 on UI)
        min_idx = row_2[0].number_input(
            "Min. Entry", min_value=1, value=1, disabled=bool(indices)
        )

        # Number of entries to extract starting from min
        max_idx = row_2[1].number_input(
            "Max. Entries", min_value=1, value=10, disabled=bool(indices)
        )

        # Return range to extract
        return int(min_idx) - 1, int(max_idx), indices, id_keys


# Extract FASTA from file's data
def extract_fasta_from_data(
    content: TextIOWrapper,
    min_idx: int = None,
    max_idx: int = None,
    indices: list = [],
    id_keys: list = [],
) -> pd.DataFrame:

    # Prepare empty dictionary to store extracted data
    data = {k: [] for k in FASTA_COLUMNS}

    """
    Remember that according to the FASTA standard, 
    each 2 subsequent lines correspond to one sequence data.
    """
    # Get FASTA by index
    if indices:
        # Lambda function returning tuple of pair indices
        make_pair = lambda j: [j, j + 1]

        # Entries are indexed starting from 1 on UI
        idx = np.ravel([make_pair((i - 1) * 2) for i in indices])

        # Get an iterator
        it = iter(itemgetter(*idx)(content.readlines()))

    # Get FASTA by index range (slices)
    else:
        # Return the entry index limits for subsequent lines
        min_idx = min_idx * 2 if min_idx else None
        max_idx = max_idx * 2 if max_idx else None

        # Get an iterator
        it = iter(content.readlines()[min_idx:max_idx])

    # Flag- Empty content
    is_empty = True

    # Flag - FASTA compliant
    is_fasta = False

    # Flag - Searching by ID
    search_by_id = bool(id_keys)

    for line in it:
        if line:
            # At least one non-empty line - disable empty content flag
            is_empty = False

            if line.startswith(">"):
                # At least one properly formatted line - enable FASTA compliant flag
                is_fasta = True

                # Sequence is below the single-line description
                sequence = next(it).strip()

                # Split descriptors into ID and remaining description
                descriptors = re.split(
                    pattern=r",\s+|\s+(?!.*?,\s+)", string=line[1:].strip(), maxsplit=1
                )

                # Descriptors expect to return more than just ID
                if len(descriptors) == 2:
                    id, description = descriptors

                # Set description to empty string if not found
                else:
                    id, description = (descriptors[0], "")

                # Flag - add entry, dependent on Searching by ID flag
                add_entry = True

                # Search by ID if keys parameeter was provided
                if search_by_id:

                    # Return the data if no more IDs are left to match
                    if not id_keys:
                        return pd.DataFrame(data)

                    # Disable before searching by ID
                    add_entry = False
                    for i, key in enumerate(id_keys):
                        if key == id:
                            # Mark the entry for adding if ID matches the key
                            add_entry = True

                            # Delete matched ID from keys
                            del id_keys[i]
                            break

                # Add entry
                if add_entry:
                    data["ID"].append(id)
                    data["DESCRIPTION"].append(description)
                    data["SEQUENCE"].append(sequence)

    # Raise FASTAFormattingError if fiile contains FASTA non-compliant content
    if not (is_empty or is_fasta):
        raise FASTAFormattingError

    # Convert dictionary to a DataFrame
    return pd.DataFrame(data)
