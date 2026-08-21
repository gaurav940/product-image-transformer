import io
import os
import tempfile
import zipfile

import streamlit as st

from image_transformation import (
    process_excel,
    DEFAULT_SKU_COLUMN,
    MIN_W,
    MIN_H,
    TARGET_RATIO,
    TARGET_W_H_RATIO,
)


# ============================================================
# PAGE CONFIG
# ============================================================

st.set_page_config(
    page_title="Product Image Transformer",
    page_icon="🖼️",
    layout="wide",
)


# ============================================================
# SESSION STATE
# ============================================================

if "zip_data" not in st.session_state:
    st.session_state.zip_data = None

if "report_data" not in st.session_state:
    st.session_state.report_data = None

if "summary" not in st.session_state:
    st.session_state.summary = None

if "processed_filename" not in st.session_state:
    st.session_state.processed_filename = None


# ============================================================
# HEADER
# ============================================================

st.title(
    "Product Image Transformer"
)

st.markdown(
    """
Upload a catalogue containing:

**Parent Group Id | Image Url 1 | Image Url 2 ... Image Url 8**

Every image will be downloaded, validated and transformed to the
platform image ratio **without cropping or stretching the original image**.
"""
)


# ============================================================
# RULES
# ============================================================

with st.expander(
    "Transformation rules",
    expanded=False,
):

    col1, col2, col3 = st.columns(
        3
    )

    with col1:

        st.markdown(
            f"""
### Minimum dimensions

- Width ≥ **{MIN_W}px**
- Height ≥ **{MIN_H}px**
"""
        )

    with col2:

        st.markdown(
            f"""
### Target ratio

- W:H = **{MIN_W}:{MIN_H}**
- W:H ≈ **{TARGET_W_H_RATIO:.4f}**
- H:W ≈ **{TARGET_RATIO:.4f}**
"""
        )

    with col3:

        st.markdown(
            """
### Transformation

- No cropping
- No stretching
- Edge-generated canvas
- Up to 8 images / SKU
"""
        )


# ============================================================
# UPLOAD
# ============================================================

st.subheader(
    "1. Upload catalogue"
)

uploaded_file = st.file_uploader(
    "Upload Excel catalogue",
    type=[
        "xlsx"
    ],
    help=(
        "Expected headers include "
        "'Parent Group Id', "
        "'Image Url 1', "
        "'Image Url 2', etc."
    ),
)


# ============================================================
# SETTINGS
# ============================================================

with st.expander(
    "Advanced settings",
    expanded=False,
):

    sku_column = st.text_input(
        "Parent SKU column",
        value=DEFAULT_SKU_COLUMN,
    )

    sheet_name = st.text_input(
        "Excel sheet name",
        value="",
        help=(
            "Leave blank to use the first sheet."
        ),
    )


# ============================================================
# FILE LOADED
# ============================================================

if uploaded_file is not None:

    st.success(
        f"Loaded: {uploaded_file.name}"
    )


# ============================================================
# PROCESS BUTTON
# ============================================================

process_clicked = st.button(
    "Process Images",
    type="primary",
    use_container_width=True,
    disabled=(
        uploaded_file is None
    ),
)


# ============================================================
# PROCESS
# ============================================================

if (
    process_clicked
    and uploaded_file is not None
):

    # Clear old outputs.

    st.session_state.zip_data = None
    st.session_state.report_data = None
    st.session_state.summary = None
    st.session_state.processed_filename = None

    st.divider()

    st.subheader(
        "2. Processing Images"
    )

    # --------------------------------------------------------
    # PROGRESS UI
    # --------------------------------------------------------

    progress_bar = st.progress(
        0.0,
        text="Preparing catalogue..."
    )

    current_image_box = st.empty()

    current_status_box = st.empty()

    progress_text_box = st.empty()

    # Live counters.

    counter_container = st.container()

    counter_cols = counter_container.columns(
        4
    )

    completed_metric = (
        counter_cols[0].empty()
    )

    success_metric = (
        counter_cols[1].empty()
    )

    blocked_metric = (
        counter_cols[2].empty()
    )

    error_metric = (
        counter_cols[3].empty()
    )

    live_stats = {
        "completed": 0,
        "success": 0,
        "blocked": 0,
        "errors": 0,
    }

    # --------------------------------------------------------
    # METRIC RENDERER
    # --------------------------------------------------------

    def render_live_metrics():

        completed_metric.metric(
            "Completed",
            live_stats[
                "completed"
            ],
        )

        success_metric.metric(
            "Successful",
            live_stats[
                "success"
            ],
        )

        blocked_metric.metric(
            "Blocked",
            live_stats[
                "blocked"
            ],
        )

        error_metric.metric(
            "Errors",
            live_stats[
                "errors"
            ],
        )

    render_live_metrics()

    # --------------------------------------------------------
    # CALLBACK CALLED FROM image_transformation.py
    # --------------------------------------------------------

    def update_progress(
        completed,
        total,
        parent_group_id,
        image_position,
        status,
    ):

        if total <= 0:
            return

        # ----------------------------------------------------
        # CURRENT IMAGE
        # ----------------------------------------------------

        if status == "PROCESSING":

            current_image_box.info(
                f"Processing "
                f"**{parent_group_id}** "
                f"· Image {image_position}"
            )

            current_status_box.write(
                "Downloading and transforming..."
            )

            progress_text_box.write(
                f"{completed} of "
                f"{total} images completed"
            )

            return

        # ----------------------------------------------------
        # COMPLETED IMAGE
        # ----------------------------------------------------

        live_stats[
            "completed"
        ] = completed

        if status.startswith(
            "SUCCESS"
        ):

            live_stats[
                "success"
            ] += 1

            current_status_box.success(
                status
            )

        elif status.startswith(
            "BLOCKED"
        ):

            live_stats[
                "blocked"
            ] += 1

            current_status_box.warning(
                status
            )

        else:

            live_stats[
                "errors"
            ] += 1

            current_status_box.error(
                status
            )

        progress = (
            completed / total
        )

        progress_bar.progress(
            progress,
            text=(
                f"{completed} / "
                f"{total} images processed"
            ),
        )

        current_image_box.info(
            f"**{parent_group_id}** "
            f"· Image {image_position}"
        )

        progress_text_box.write(
            f"{progress * 100:.0f}% complete"
        )

        render_live_metrics()

    # --------------------------------------------------------
    # RUN JOB
    # --------------------------------------------------------

    try:

        with tempfile.TemporaryDirectory() as temp_dir:

            input_path = os.path.join(
                temp_dir,
                "input.xlsx",
            )

            images_dir = os.path.join(
                temp_dir,
                "processed_images",
            )

            report_path = os.path.join(
                temp_dir,
                "processing_report.xlsx",
            )

            os.makedirs(
                images_dir,
                exist_ok=True,
            )

            # ------------------------------------------------
            # SAVE UPLOADED FILE
            # ------------------------------------------------

            with open(
                input_path,
                "wb",
            ) as file:

                file.write(
                    uploaded_file.getbuffer()
                )

            # ------------------------------------------------
            # PROCESS EXCEL
            # ------------------------------------------------

            summary = process_excel(
                input_path=input_path,
                output_dir=images_dir,
                report_path=report_path,
                sku_column=(
                    sku_column.strip()
                    or DEFAULT_SKU_COLUMN
                ),
                sheet_name=(
                    sheet_name.strip()
                    if sheet_name.strip()
                    else None
                ),
                progress_callback=(
                    update_progress
                ),
            )

            # ------------------------------------------------
            # REPORT BYTES
            # ------------------------------------------------

            with open(
                report_path,
                "rb",
            ) as file:

                report_data = (
                    file.read()
                )

            # ------------------------------------------------
            # CREATE ZIP
            # ------------------------------------------------

            zip_buffer = io.BytesIO()

            with zipfile.ZipFile(
                zip_buffer,
                mode="w",
                compression=(
                    zipfile.ZIP_DEFLATED
                ),
            ) as zip_file:

                # Images
                for filename in sorted(
                    os.listdir(
                        images_dir
                    )
                ):

                    file_path = os.path.join(
                        images_dir,
                        filename,
                    )

                    if not os.path.isfile(
                        file_path
                    ):
                        continue

                    zip_file.write(
                        file_path,
                        arcname=os.path.join(
                            "processed_images",
                            filename,
                        ),
                    )

                # Report
                zip_file.write(
                    report_path,
                    arcname=(
                        "processing_report.xlsx"
                    ),
                )

            zip_buffer.seek(
                0
            )

            # ------------------------------------------------
            # SAVE OUTPUT IN SESSION
            # ------------------------------------------------

            st.session_state.zip_data = (
                zip_buffer.getvalue()
            )

            st.session_state.report_data = (
                report_data
            )

            st.session_state.summary = (
                summary
            )

            st.session_state.processed_filename = (
                uploaded_file.name
            )

            # ------------------------------------------------
            # FINISHED
            # ------------------------------------------------

            progress_bar.progress(
                1.0,
                text=(
                    f"{summary['total']} / "
                    f"{summary['total']} "
                    f"images processed"
                ),
            )

            current_image_box.empty()

            current_status_box.success(
                "Processing complete."
            )

            progress_text_box.write(
                "100% complete"
            )

    except Exception as error:

        progress_bar.empty()

        current_image_box.empty()

        current_status_box.error(
            "Processing failed."
        )

        st.exception(
            error
        )


# ============================================================
# FINAL SUMMARY
# ============================================================

if st.session_state.summary:

    summary = (
        st.session_state.summary
    )

    st.divider()

    st.subheader(
        "3. Processing Summary"
    )

    col1, col2, col3, col4, col5 = (
        st.columns(
            5
        )
    )

    col1.metric(
        "Images",
        summary[
            "total"
        ],
    )

    col2.metric(
        "Successful",
        summary[
            "successful"
        ],
    )

    col3.metric(
        "Transformed",
        summary[
            "transformed"
        ],
    )

    col4.metric(
        "Blocked",
        summary[
            "blocked"
        ],
    )

    col5.metric(
        "Errors",
        summary[
            "errors"
        ],
    )

    # --------------------------------------------------------
    # DETAILED STATUS BREAKDOWN
    # --------------------------------------------------------

    with st.expander(
        "Status breakdown"
    ):

        for status, count in sorted(
            summary[
                "statuses"
            ].items()
        ):

            st.write(
                f"**{status}:** "
                f"{count}"
            )


# ============================================================
# DOWNLOADS
# ============================================================

if (
    st.session_state.zip_data
    is not None
):

    st.divider()

    st.subheader(
        "4. Download Results"
    )

    st.markdown(
        """
The ZIP contains:

- All successfully processed JPG images
- `processing_report.xlsx`
"""
    )

    col1, col2 = st.columns(
        2
    )

    with col1:

        st.download_button(
            label=(
                "Download Images + Report"
            ),
            data=(
                st.session_state.zip_data
            ),
            file_name=(
                "image_processing_output.zip"
            ),
            mime="application/zip",
            type="primary",
            use_container_width=True,
        )

    with col2:

        st.download_button(
            label=(
                "Download Processing Report"
            ),
            data=(
                st.session_state.report_data
            ),
            file_name=(
                "processing_report.xlsx"
            ),
            mime=(
                "application/"
                "vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),
            use_container_width=True,
        )