#!/usr/bin/env python3

"""
image_transformation.py

Bulk product image transformation engine.

Expected Excel format:

    Parent Group Id | Image Url 1 | Image Url 2 | ... | Image Url 8

Rules:

    Minimum Width  = 660 px
    Minimum Height = 900 px

    Target W:H = 660:900 = 0.733333...
    Target H:W = 900:660 = 1.363636...

Behaviour:

    - Does NOT crop the source image.
    - Does NOT stretch/distort the source image.
    - Adds canvas where required.
    - Canvas is generated from the image edge/background.
    - Supports up to 8 images per Parent Group Id.
    - Images below 660 x 900 are blocked.
    - One failed image does NOT fail the entire SKU.
    - Generates a "Processing Results" sheet.
    - Supports a progress callback for Streamlit.

Dependencies:

    pip install pillow numpy requests openpyxl
"""

import os
import re
import sys
import time
import argparse
from io import BytesIO
from collections import Counter

import numpy as np
import requests

from PIL import Image, ImageFilter, ImageOps
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter


# ============================================================
# PLATFORM CONFIGURATION
# ============================================================

MIN_W = 660
MIN_H = 900

# H:W
TARGET_RATIO = MIN_H / MIN_W

# W:H
TARGET_W_H_RATIO = MIN_W / MIN_H

MAX_IMAGES_PER_SKU = 8

EDGE_SAMPLE_PX = 6

SEAM_BLUR_BAND = 14
SEAM_BLUR_RADIUS = 6

MAX_DOWNLOAD_MB = 30

CONNECT_TIMEOUT = 10
READ_TIMEOUT = 30

DOWNLOAD_RETRIES = 3

DEFAULT_SKU_COLUMN = "Parent Group Id"

RESULT_SHEET_NAME = "Processing Results"



def normalize_source_url(url):
    """
    Convert known share/view URLs into direct image-download URLs.
    """

    url = str(url).strip()

    # Google Drive:
    # https://drive.google.com/file/d/FILE_ID/view?usp=sharing
    drive_match = re.search(
        r"drive\.google\.com/file/d/([^/]+)",
        url,
    )

    if drive_match:
        file_id = drive_match.group(1)

        return (
            "https://drive.usercontent.google.com/"
            f"download?id={file_id}&export=download"
        )

    return url


# ============================================================
# IMAGE NORMALIZATION
# ===========================================================

def normalize_image(img):
    """
    Apply EXIF orientation and convert the image to RGB.

    Transparent images are placed on a white background.
    """

    img = ImageOps.exif_transpose(img)

    if img.mode in ("RGBA", "LA"):
        rgba = img.convert("RGBA")

        background = Image.new(
            "RGBA",
            rgba.size,
            (255, 255, 255, 255),
        )

        background.alpha_composite(rgba)

        return background.convert("RGB")

    if img.mode == "P" and "transparency" in img.info:
        rgba = img.convert("RGBA")

        background = Image.new(
            "RGBA",
            rgba.size,
            (255, 255, 255, 255),
        )

        background.alpha_composite(rgba)

        return background.convert("RGB")

    return img.convert("RGB")


# ============================================================
# BACKGROUND / NOISE GENERATION
# ============================================================


def estimate_local_noise(strip):
    """
    Estimate local image grain/noise using adjacent pixel differences.
    """

    strip = strip.astype(np.float32)

    differences = []

    if strip.shape[0] > 1:
        differences.append(
            np.diff(strip, axis=0).ravel()
        )

    if strip.shape[1] > 1:
        differences.append(
            np.diff(strip, axis=1).ravel()
        )

    if not differences:
        return 0.8

    values = np.concatenate(differences)

    std = values.std() / np.sqrt(2)

    return float(
        np.clip(
            std,
            0.4,
            3.0,
        )
    )


def build_edge_fill(
    edge_strip,
    pad_size,
    axis,
):
    """
    Generate a padded area from the true image edge.

    axis="v"
        Top/bottom padding.

    axis="h"
        Left/right padding.
    """

    if pad_size <= 0:
        raise ValueError(
            "pad_size must be greater than zero"
        )

    if axis == "v":

        profile = edge_strip.mean(
            axis=0,
            keepdims=True,
        )

        return np.repeat(
            profile,
            pad_size,
            axis=0,
        )

    if axis == "h":

        profile = edge_strip.mean(
            axis=1,
            keepdims=True,
        )

        return np.repeat(
            profile,
            pad_size,
            axis=1,
        )

    raise ValueError(
        "axis must be 'v' or 'h'"
    )


def add_matched_noise(
    region,
    noise_std,
    seed,
):
    """
    Add subtle grain to the generated canvas.
    """

    rng = np.random.default_rng(seed)

    noise = rng.normal(
        0,
        noise_std,
        size=region.shape,
    )

    output = (
        region.astype(np.float32)
        + noise
    )

    return np.clip(
        output,
        0,
        255,
    )


# ============================================================
# VALIDATION
# ============================================================


def validate_resolution(
    width,
    height,
):
    """
    Enforce minimum supported resolution.
    """

    if width < MIN_W or height < MIN_H:

        raise ValueError(
            f"Image {width}x{height} is below "
            f"minimum supported resolution "
            f"{MIN_W}x{MIN_H}"
        )


# ============================================================
# IMAGE TRANSFORMATION
# ============================================================


def pad_to_target_ratio(img):
    """
    Pad image to H:W = 900:660.

    No cropping.
    No stretching.
    No distortion.

    Returns:

        output_image
        transformation_description
    """

    img = normalize_image(img)

    width, height = img.size

    validate_resolution(
        width,
        height,
    )

    current_ratio = height / width

    # --------------------------------------------------------
    # ALREADY AT TARGET
    # --------------------------------------------------------

    if abs(current_ratio - TARGET_RATIO) < 0.001:

        return (
            img,
            "NONE",
        )

    arr = np.array(
        img,
        dtype=np.float32,
    )

    # ========================================================
    # IMAGE TOO WIDE / SHORT
    #
    # Need additional HEIGHT.
    # Add top + bottom padding.
    # ========================================================

    if current_ratio < TARGET_RATIO:

        new_width = width

        new_height = round(
            width * TARGET_RATIO
        )

        pad_total = (
            new_height - height
        )

        pad_top = (
            pad_total // 2
        )

        pad_bottom = (
            pad_total - pad_top
        )

        canvas = np.zeros(
            (
                new_height,
                new_width,
                3,
            ),
            dtype=np.float32,
        )

        # Original image
        canvas[
            pad_top:pad_top + height,
            :,
            :
        ] = arr

        edge_size = min(
            EDGE_SAMPLE_PX,
            height,
        )

        top_strip = arr[
            :edge_size,
            :,
            :
        ]

        bottom_strip = arr[
            -edge_size:,
            :,
            :
        ]

        if pad_top > 0:

            top_fill = build_edge_fill(
                top_strip,
                pad_top,
                axis="v",
            )

            top_fill = add_matched_noise(
                top_fill,
                estimate_local_noise(
                    top_strip
                ),
                seed=1,
            )

            canvas[
                :pad_top,
                :,
                :
            ] = top_fill

        if pad_bottom > 0:

            bottom_fill = build_edge_fill(
                bottom_strip,
                pad_bottom,
                axis="v",
            )

            bottom_fill = add_matched_noise(
                bottom_fill,
                estimate_local_noise(
                    bottom_strip
                ),
                seed=2,
            )

            canvas[
                pad_top + height:,
                :,
                :
            ] = bottom_fill

        seams = []

        if pad_top > 0:
            seams.append(
                pad_top
            )

        if pad_bottom > 0:
            seams.append(
                pad_top + height
            )

        vertical_padding = True

        transformation = (
            f"PAD_TOP_BOTTOM | "
            f"top={pad_top}px | "
            f"bottom={pad_bottom}px"
        )

    # ========================================================
    # IMAGE TOO TALL / NARROW
    #
    # Need additional WIDTH.
    # Add left + right padding.
    # ========================================================

    else:

        new_height = height

        new_width = round(
            height / TARGET_RATIO
        )

        pad_total = (
            new_width - width
        )

        pad_left = (
            pad_total // 2
        )

        pad_right = (
            pad_total - pad_left
        )

        canvas = np.zeros(
            (
                new_height,
                new_width,
                3,
            ),
            dtype=np.float32,
        )

        # Original image
        canvas[
            :,
            pad_left:pad_left + width,
            :
        ] = arr

        edge_size = min(
            EDGE_SAMPLE_PX,
            width,
        )

        left_strip = arr[
            :,
            :edge_size,
            :
        ]

        right_strip = arr[
            :,
            -edge_size:,
            :
        ]

        if pad_left > 0:

            left_fill = build_edge_fill(
                left_strip,
                pad_left,
                axis="h",
            )

            left_fill = add_matched_noise(
                left_fill,
                estimate_local_noise(
                    left_strip
                ),
                seed=3,
            )

            canvas[
                :,
                :pad_left,
                :
            ] = left_fill

        if pad_right > 0:

            right_fill = build_edge_fill(
                right_strip,
                pad_right,
                axis="h",
            )

            right_fill = add_matched_noise(
                right_fill,
                estimate_local_noise(
                    right_strip
                ),
                seed=4,
            )

            canvas[
                :,
                pad_left + width:,
                :
            ] = right_fill

        seams = []

        if pad_left > 0:
            seams.append(
                pad_left
            )

        if pad_right > 0:
            seams.append(
                pad_left + width
            )

        vertical_padding = False

        transformation = (
            f"PAD_LEFT_RIGHT | "
            f"left={pad_left}px | "
            f"right={pad_right}px"
        )

    # ========================================================
    # CONVERT BACK TO PIL
    # ========================================================

    output_image = Image.fromarray(
        np.clip(
            canvas,
            0,
            255,
        ).astype(
            np.uint8
        )
    )

    # ========================================================
    # SOFTEN SEAMS
    # ========================================================

    for seam in seams:

        if vertical_padding:

            box = (
                0,
                max(
                    0,
                    seam - SEAM_BLUR_BAND,
                ),
                new_width,
                min(
                    new_height,
                    seam + SEAM_BLUR_BAND,
                ),
            )

        else:

            box = (
                max(
                    0,
                    seam - SEAM_BLUR_BAND,
                ),
                0,
                min(
                    new_width,
                    seam + SEAM_BLUR_BAND,
                ),
                new_height,
            )

        crop = output_image.crop(
            box
        )

        crop = crop.filter(
            ImageFilter.GaussianBlur(
                radius=SEAM_BLUR_RADIUS
            )
        )

        output_image.paste(
            crop,
            box,
        )

    return (
        output_image,
        transformation,
    )


# ============================================================
# IMAGE DOWNLOAD
# ============================================================


def download_image(url):
    """
    Download image with retries, redirects and timeout handling.
    """
    url = normalize_source_url(url)
    headers = {
        "User-Agent": (
            "Mozilla/5.0 "
            "(Macintosh; Intel Mac OS X 10_15_7) "
            "AppleWebKit/537.36 "
            "(KHTML, like Gecko) "
            "Chrome/151.0 Safari/537.36"
        ),
        "Accept": (
            "image/avif,"
            "image/webp,"
            "image/apng,"
            "image/*,"
            "*/*;q=0.8"
        ),
        "Accept-Language": (
            "en-US,en;q=0.9"
        ),
    }

    max_bytes = (
        MAX_DOWNLOAD_MB
        * 1024
        * 1024
    )

    last_error = None

    for attempt in range(
        1,
        DOWNLOAD_RETRIES + 1,
    ):

        try:

            with requests.get(
                url,
                headers=headers,
                timeout=(
                    CONNECT_TIMEOUT,
                    READ_TIMEOUT,
                ),
                stream=True,
                allow_redirects=True,
            ) as response:

                if response.status_code in (
                    429,
                    500,
                    502,
                    503,
                    504,
                ):

                    raise requests.exceptions.HTTPError(
                        f"HTTP {response.status_code}",
                        response=response,
                    )

                response.raise_for_status()

                content_type = (
                    response.headers
                    .get(
                        "Content-Type",
                        ""
                    )
                    .lower()
                )

                if (
                    content_type
                    and "image/" not in content_type
                    and "application/octet-stream"
                    not in content_type
                ):

                    raise ValueError(
                        "URL did not return an image. "
                        f"Content-Type={content_type}"
                    )

                content_length = (
                    response.headers.get(
                        "Content-Length"
                    )
                )

                if content_length:

                    try:
                        content_length_int = int(
                            content_length
                        )

                        if (
                            content_length_int
                            > max_bytes
                        ):
                            raise ValueError(
                                f"Image is larger than "
                                f"{MAX_DOWNLOAD_MB} MB"
                            )

                    except ValueError as error:

                        if (
                            "larger than"
                            in str(error)
                        ):
                            raise

                buffer = BytesIO()

                downloaded = 0

                for chunk in response.iter_content(
                    chunk_size=64 * 1024
                ):

                    if not chunk:
                        continue

                    downloaded += len(
                        chunk
                    )

                    if downloaded > max_bytes:

                        raise ValueError(
                            f"Image is larger than "
                            f"{MAX_DOWNLOAD_MB} MB"
                        )

                    buffer.write(
                        chunk
                    )

            buffer.seek(0)

            image = Image.open(
                buffer
            )

            # Decode immediately.
            image.load()

            return image

        except (
            requests.exceptions.Timeout,
            requests.exceptions.ConnectionError,
            requests.exceptions.HTTPError,
        ) as error:

            last_error = error

            status_code = None

            if (
                isinstance(
                    error,
                    requests.exceptions.HTTPError,
                )
                and error.response is not None
            ):
                status_code = (
                    error.response.status_code
                )

            # Permanent HTTP failures should not be retried.
            if (
                status_code is not None
                and status_code not in (
                    429,
                    500,
                    502,
                    503,
                    504,
                )
            ):
                raise

            if attempt < DOWNLOAD_RETRIES:

                time.sleep(
                    attempt
                )

    raise last_error


# ============================================================
# HELPERS
# ============================================================


def safe_filename(value):
    """
    Convert Parent Group Id into a safe filesystem filename.
    """

    value = str(
        value
    ).strip()

    value = re.sub(
        r'[<>:"/\\|?*]',
        "_",
        value,
    )

    value = re.sub(
        r"\s+",
        "_",
        value,
    )

    value = value.strip(
        "._"
    )

    if not value:
        return "unknown"

    return value


def get_image_source(cell):
    """
    Read URL from a normal Excel cell or Excel hyperlink.
    """

    if cell.hyperlink:

        target = (
            cell.hyperlink.target
        )

        if target:
            return str(
                target
            ).strip()

    if cell.value is None:
        return None

    value = str(
        cell.value
    ).strip()

    if not value:
        return None

    return value


# ============================================================
# EXCEL HEADER DETECTION
# ============================================================


def find_header_column(
    worksheet,
    expected_header,
):
    """
    Find a header in row 1, case-insensitive.
    """

    expected = (
        expected_header
        .strip()
        .lower()
    )

    for cell in worksheet[1]:

        if cell.value is None:
            continue

        actual = str(
            cell.value
        ).strip().lower()

        if actual == expected:
            return cell.column

    return None


def find_image_columns(
    worksheet,
):
    """
    Detect:

        Image Url 1
        Image Url 2
        ...
        Image Url 8

    Case-insensitive.
    """

    pattern = re.compile(
        r"^image\s*url\s*(\d+)$",
        re.IGNORECASE,
    )

    image_columns = []

    for cell in worksheet[1]:

        if cell.value is None:
            continue

        header = str(
            cell.value
        ).strip()

        match = pattern.match(
            header
        )

        if not match:
            continue

        position = int(
            match.group(1)
        )

        if not (
            1 <= position <= MAX_IMAGES_PER_SKU
        ):
            continue

        image_columns.append(
            {
                "position": position,
                "column": cell.column,
                "header": header,
            }
        )

    image_columns.sort(
        key=lambda item: item[
            "position"
        ]
    )

    return image_columns


# ============================================================
# PROCESS ONE IMAGE
# ============================================================


def process_image(
    source_url,
    parent_group_id,
    image_position,
    output_dir,
):
    """
    Download, validate, transform and save one image.
    """

    result = {
        "parent_group_id": str(
            parent_group_id
        ),
        "image_position": image_position,
        "source_url": source_url,
        "processing_status": "",
        "original_dimensions": "",
        "original_w_h_ratio": "",
        "output_dimensions": "",
        "output_w_h_ratio": "",
        "transformation": "",
        "output_file": "",
        "processing_error": "",
    }

    try:

        # ----------------------------------------------------
        # DOWNLOAD
        # ----------------------------------------------------

        image = download_image(
            source_url
        )

        image = normalize_image(
            image
        )

        original_width, original_height = (
            image.size
        )

        result[
            "original_dimensions"
        ] = (
            f"{original_width}x"
            f"{original_height}"
        )

        result[
            "original_w_h_ratio"
        ] = round(
            original_width
            / original_height,
            4,
        )

        # ----------------------------------------------------
        # MINIMUM RESOLUTION CHECK
        # ----------------------------------------------------

        if (
            original_width < MIN_W
            or original_height < MIN_H
        ):

            result[
                "processing_status"
            ] = (
                "BLOCKED_MIN_RESOLUTION"
            )

            result[
                "processing_error"
            ] = (
                f"Image "
                f"{original_width}x"
                f"{original_height} "
                f"is below minimum "
                f"{MIN_W}x{MIN_H}"
            )

            return result

        # ----------------------------------------------------
        # TRANSFORM
        # ----------------------------------------------------

        transformed, transformation = (
            pad_to_target_ratio(
                image
            )
        )

        output_width, output_height = (
            transformed.size
        )

        result[
            "output_dimensions"
        ] = (
            f"{output_width}x"
            f"{output_height}"
        )

        result[
            "output_w_h_ratio"
        ] = round(
            output_width
            / output_height,
            4,
        )

        result[
            "transformation"
        ] = transformation

        # ----------------------------------------------------
        # SAVE
        #
        # SKU123_01.jpg
        # SKU123_02.jpg
        # ----------------------------------------------------

        filename = (
            f"{safe_filename(parent_group_id)}_"
            f"{image_position:02d}.jpg"
        )

        output_path = os.path.join(
            output_dir,
            filename,
        )

        transformed.save(
            output_path,
            format="JPEG",
            quality=95,
            optimize=True,
            subsampling=0,
        )

        result[
            "output_file"
        ] = filename

        if transformation == "NONE":

            result[
                "processing_status"
            ] = (
                "SUCCESS_NO_CHANGE"
            )

        else:

            result[
                "processing_status"
            ] = (
                "SUCCESS_TRANSFORMED"
            )

        return result

    except requests.exceptions.Timeout:

        result[
            "processing_status"
        ] = (
            "ERROR_TIMEOUT"
        )

        result[
            "processing_error"
        ] = (
            "Image download timed out"
        )

        return result

    except requests.exceptions.HTTPError as error:

        status_code = ""

        if error.response is not None:

            status_code = (
                error.response.status_code
            )

        result[
            "processing_status"
        ] = (
            "ERROR_HTTP"
        )

        result[
            "processing_error"
        ] = (
            f"HTTP {status_code}: "
            f"{error}"
        )

        return result

    except requests.exceptions.ConnectionError as error:

        result[
            "processing_status"
        ] = (
            "ERROR_CONNECTION"
        )

        result[
            "processing_error"
        ] = str(
            error
        )

        return result

    except Exception as error:

        result[
            "processing_status"
        ] = (
            "ERROR"
        )

        result[
            "processing_error"
        ] = str(
            error
        )

        return result


# ============================================================
# RESULTS SHEET
# ============================================================


RESULT_HEADERS = [
    "Parent Group Id",
    "Image Position",
    "Source Column",
    "Source URL",
    "Processing Status",
    "Original Dimensions",
    "Original W:H Ratio",
    "Output Dimensions",
    "Output W:H Ratio",
    "Transformation",
    "Output File",
    "Processing Error",
]


def prepare_results_sheet(
    workbook,
):
    """
    Replace previous results sheet, if one exists.
    """

    if RESULT_SHEET_NAME in workbook.sheetnames:

        existing_sheet = workbook[
            RESULT_SHEET_NAME
        ]

        workbook.remove(
            existing_sheet
        )

    worksheet = workbook.create_sheet(
        RESULT_SHEET_NAME
    )

    for column_index, header in enumerate(
        RESULT_HEADERS,
        start=1,
    ):

        cell = worksheet.cell(
            row=1,
            column=column_index,
            value=header,
        )

        cell.font = Font(
            bold=True,
            color="FFFFFF",
        )

        cell.fill = PatternFill(
            fill_type="solid",
            fgColor="1F4E78",
        )

        cell.alignment = Alignment(
            vertical="center",
        )

    worksheet.freeze_panes = "A2"

    return worksheet


def write_result_row(
    worksheet,
    source_column,
    result,
):
    """
    Add one processing result to the results sheet.
    """

    worksheet.append(
        [
            result[
                "parent_group_id"
            ],
            result[
                "image_position"
            ],
            source_column,
            result[
                "source_url"
            ],
            result[
                "processing_status"
            ],
            result[
                "original_dimensions"
            ],
            result[
                "original_w_h_ratio"
            ],
            result[
                "output_dimensions"
            ],
            result[
                "output_w_h_ratio"
            ],
            result[
                "transformation"
            ],
            result[
                "output_file"
            ],
            result[
                "processing_error"
            ],
        ]
    )


def format_results_sheet(
    worksheet,
):
    """
    Make the results sheet easier to review.
    """

    worksheet.auto_filter.ref = (
        worksheet.dimensions
    )

    widths = {
        1: 30,
        2: 15,
        3: 18,
        4: 75,
        5: 28,
        6: 22,
        7: 20,
        8: 22,
        9: 20,
        10: 42,
        11: 42,
        12: 60,
    }

    for column_index, width in widths.items():

        worksheet.column_dimensions[
            get_column_letter(
                column_index
            )
        ].width = width

    success_fill = PatternFill(
        fill_type="solid",
        fgColor="E2F0D9",
    )

    warning_fill = PatternFill(
        fill_type="solid",
        fgColor="FCE5CD",
    )

    error_fill = PatternFill(
        fill_type="solid",
        fgColor="F4CCCC",
    )

    # Processing Status = column 5
    for row_index in range(
        2,
        worksheet.max_row + 1,
    ):

        status_cell = worksheet.cell(
            row=row_index,
            column=5,
        )

        status = str(
            status_cell.value or ""
        )

        if status.startswith(
            "SUCCESS"
        ):

            status_cell.fill = (
                success_fill
            )

        elif status.startswith(
            "BLOCKED"
        ):

            status_cell.fill = (
                warning_fill
            )

        elif status.startswith(
            "ERROR"
        ):

            status_cell.fill = (
                error_fill
            )


# ============================================================
# EXCEL PROCESSOR
# ============================================================


def process_excel(
    input_path,
    output_dir,
    report_path,
    sku_column=DEFAULT_SKU_COLUMN,
    sheet_name=None,
    progress_callback=None,
):
    """
    Process complete Excel catalogue.

    progress_callback, when supplied, receives:

        progress_callback(
            completed,
            total,
            parent_group_id,
            image_position,
            status,
        )

    Status will first be:

        PROCESSING

    and then:

        SUCCESS_TRANSFORMED
        SUCCESS_NO_CHANGE
        BLOCKED_MIN_RESOLUTION
        ERROR_HTTP
        etc.

    Returns a summary dictionary.
    """

    workbook = load_workbook(
        input_path
    )

    # --------------------------------------------------------
    # INPUT WORKSHEET
    # --------------------------------------------------------

    if sheet_name:

        if sheet_name not in workbook.sheetnames:

            raise ValueError(
                f"Sheet '{sheet_name}' not found. "
                f"Available sheets: "
                f"{workbook.sheetnames}"
            )

        worksheet = workbook[
            sheet_name
        ]

    else:

        possible_sheets = [
            sheet
            for sheet in workbook.worksheets
            if sheet.title != RESULT_SHEET_NAME
        ]

        if not possible_sheets:

            raise ValueError(
                "No input worksheet found."
            )

        worksheet = possible_sheets[
            0
        ]

    # --------------------------------------------------------
    # PARENT GROUP ID COLUMN
    # --------------------------------------------------------

    sku_column_index = find_header_column(
        worksheet,
        sku_column,
    )

    if sku_column_index is None:

        headers = [
            str(cell.value)
            for cell in worksheet[1]
            if cell.value is not None
        ]

        raise ValueError(
            f"Could not find "
            f"'{sku_column}'. "
            f"Available headers: "
            f"{headers}"
        )

    # --------------------------------------------------------
    # IMAGE URL COLUMNS
    # --------------------------------------------------------

    image_columns = find_image_columns(
        worksheet
    )

    if not image_columns:

        raise ValueError(
            "No image columns found. "
            "Expected 'Image Url 1', "
            "'Image Url 2', etc."
        )

    # --------------------------------------------------------
    # OUTPUT
    # --------------------------------------------------------

    os.makedirs(
        output_dir,
        exist_ok=True,
    )

    results_sheet = prepare_results_sheet(
        workbook
    )

    # --------------------------------------------------------
    # COUNT TOTAL IMAGES
    # --------------------------------------------------------

    total_images = 0

    for row_index in range(
        2,
        worksheet.max_row + 1,
    ):

        parent_group_id = worksheet.cell(
            row=row_index,
            column=sku_column_index,
        ).value

        if (
            parent_group_id is None
            or not str(
                parent_group_id
            ).strip()
        ):
            continue

        for item in image_columns:

            image_cell = worksheet.cell(
                row=row_index,
                column=item[
                    "column"
                ],
            )

            if get_image_source(
                image_cell
            ):
                total_images += 1

    if total_images == 0:

        raise ValueError(
            "No image URLs were found in the workbook."
        )

    print()
    print(
        f"Sheet: {worksheet.title}"
    )

    print(
        f"Images to process: "
        f"{total_images}"
    )

    print(
        f"Minimum: "
        f"{MIN_W}x{MIN_H}"
    )

    print(
        f"Target W:H: "
        f"{TARGET_W_H_RATIO:.6f}"
    )

    print()

    # --------------------------------------------------------
    # PROCESS IMAGES
    # --------------------------------------------------------

    status_counts = Counter()

    processed_count = 0

    for row_index in range(
        2,
        worksheet.max_row + 1,
    ):

        parent_group_id = worksheet.cell(
            row=row_index,
            column=sku_column_index,
        ).value

        if (
            parent_group_id is None
            or not str(
                parent_group_id
            ).strip()
        ):
            continue

        parent_group_id = str(
            parent_group_id
        ).strip()

        for item in image_columns:

            image_cell = worksheet.cell(
                row=row_index,
                column=item[
                    "column"
                ],
            )

            source_url = get_image_source(
                image_cell
            )

            if not source_url:
                continue

            # ----------------------------------------------
            # TELL STREAMLIT WE ARE STARTING THIS IMAGE
            # ----------------------------------------------

            if progress_callback:

                progress_callback(
                    processed_count,
                    total_images,
                    parent_group_id,
                    item[
                        "position"
                    ],
                    "PROCESSING",
                )

            # ----------------------------------------------
            # PROCESS IMAGE
            # ----------------------------------------------

            result = process_image(
                source_url=source_url,
                parent_group_id=parent_group_id,
                image_position=item[
                    "position"
                ],
                output_dir=output_dir,
            )

            processed_count += 1

            status = result[
                "processing_status"
            ]

            status_counts[
                status
            ] += 1

            # ----------------------------------------------
            # WRITE RESULT
            # ----------------------------------------------

            write_result_row(
                worksheet=results_sheet,
                source_column=item[
                    "header"
                ],
                result=result,
            )

            # ----------------------------------------------
            # UPDATE STREAMLIT AFTER IMAGE COMPLETES
            # ----------------------------------------------

            if progress_callback:

                progress_callback(
                    processed_count,
                    total_images,
                    parent_group_id,
                    item[
                        "position"
                    ],
                    status,
                )

            print(
                f"["
                f"{processed_count}"
                f"/"
                f"{total_images}"
                f"] "
                f"{parent_group_id} "
                f"| Image "
                f"{item['position']} "
                f"| {status}"
            )

    # --------------------------------------------------------
    # SAVE REPORT
    # --------------------------------------------------------

    format_results_sheet(
        results_sheet
    )

    workbook.save(
        report_path
    )

    workbook.close()

    # --------------------------------------------------------
    # SUMMARY
    # --------------------------------------------------------

    successful = sum(
        count
        for status, count
        in status_counts.items()
        if status.startswith(
            "SUCCESS"
        )
    )

    transformed = status_counts.get(
        "SUCCESS_TRANSFORMED",
        0,
    )

    unchanged = status_counts.get(
        "SUCCESS_NO_CHANGE",
        0,
    )

    blocked = status_counts.get(
        "BLOCKED_MIN_RESOLUTION",
        0,
    )

    errors = sum(
        count
        for status, count
        in status_counts.items()
        if status.startswith(
            "ERROR"
        )
    )

    return {
        "total": processed_count,
        "successful": successful,
        "transformed": transformed,
        "unchanged": unchanged,
        "blocked": blocked,
        "errors": errors,
        "statuses": dict(
            status_counts
        ),
        "image_columns": [
            item["header"]
            for item in image_columns
        ],
        "input_sheet": worksheet.title,
    }


# ============================================================
# COMMAND LINE SUPPORT
# ============================================================


def main():

    parser = argparse.ArgumentParser(
        description=(
            "Process product images from an "
            "Excel catalogue containing "
            "Parent Group Id + Image Url 1..8."
        )
    )

    parser.add_argument(
        "input",
        help="Input .xlsx file",
    )

    parser.add_argument(
        "output",
        help=(
            "Folder for processed images"
        ),
    )

    parser.add_argument(
        "--sheet",
        default=None,
        help="Excel sheet name",
    )

    parser.add_argument(
        "--sku-column",
        default=DEFAULT_SKU_COLUMN,
        help=(
            "SKU column. "
            "Default: Parent Group Id"
        ),
    )

    parser.add_argument(
        "--report",
        default=None,
        help=(
            "Output report path"
        ),
    )

    args = parser.parse_args()

    input_path = os.path.abspath(
        args.input
    )

    output_dir = os.path.abspath(
        args.output
    )

    if not os.path.isfile(
        input_path
    ):

        raise FileNotFoundError(
            f"Input file not found: "
            f"{input_path}"
        )

    if not input_path.lower().endswith(
        ".xlsx"
    ):

        raise ValueError(
            "Input must be an .xlsx file."
        )

    os.makedirs(
        output_dir,
        exist_ok=True,
    )

    if args.report:

        report_path = os.path.abspath(
            args.report
        )

    else:

        report_path = os.path.join(
            output_dir,
            "processing_report.xlsx",
        )

    summary = process_excel(
        input_path=input_path,
        output_dir=output_dir,
        report_path=report_path,
        sku_column=args.sku_column,
        sheet_name=args.sheet,
    )

    print()
    print(
        "=" * 60
    )
    print(
        "PROCESSING COMPLETE"
    )
    print(
        "=" * 60
    )

    print(
        f"Total: "
        f"{summary['total']}"
    )

    print(
        f"Successful: "
        f"{summary['successful']}"
    )

    print(
        f"Transformed: "
        f"{summary['transformed']}"
    )

    print(
        f"Unchanged: "
        f"{summary['unchanged']}"
    )

    print(
        f"Blocked: "
        f"{summary['blocked']}"
    )

    print(
        f"Errors: "
        f"{summary['errors']}"
    )

    print()
    print(
        f"Report: "
        f"{report_path}"
    )


if __name__ == "__main__":

    try:

        main()

    except KeyboardInterrupt:

        print(
            "\nProcessing cancelled."
        )

        sys.exit(
            1
        )

    except Exception as error:

        print(
            f"\nERROR: {error}"
        )

        sys.exit(
            1
        )