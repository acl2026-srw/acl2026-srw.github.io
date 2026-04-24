import os
import re
import time
import getpass
import pandas as pd
import openreview
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE

# ============================================================
# CONFIG
# ============================================================

INPUT_FILE = "Decisions.xlsx"
OUTPUT_FILE = "Decisions_enriched.xlsx"
PARTIAL_OUTPUT_FILE = "Decisions_enriched_partial.xlsx"

# Save progress every N submissions
SAVE_EVERY = 10

# Your sheet must have this column
FORUM_COLUMN = "forum"

# Optional columns used only for prettier progress logs
TITLE_COLUMN = "title"
NUMBER_COLUMN = "number"


# ============================================================
# OPENREVIEW LOGIN
# ============================================================

# You can either set these environment variables:
#   export OPENREVIEW_USERNAME="your@email.com"
#   export OPENREVIEW_PASSWORD="your_password"
#
# Or the script will ask you interactively.

OPENREVIEW_USERNAME = os.environ.get("OPENREVIEW_USERNAME")
OPENREVIEW_PASSWORD = os.environ.get("OPENREVIEW_PASSWORD")

if not OPENREVIEW_USERNAME:
    OPENREVIEW_USERNAME = input("OpenReview username/email: ").strip()

if not OPENREVIEW_PASSWORD:
    OPENREVIEW_PASSWORD = getpass.getpass("OpenReview password: ")


# Try API v2 first.
# Most modern OpenReview venues use api2.openreview.net.
client = openreview.api.OpenReviewClient(
    baseurl="https://api2.openreview.net",
    username=OPENREVIEW_USERNAME,
    password=OPENREVIEW_PASSWORD,
)


# ============================================================
# FIELD CANDIDATES
# ============================================================

KEYWORD_FIELDS = [
    "keywords",
    "keyword",
    "topics",
    "topic",
    "TL;DR",
    "tl_dr",
]

CATEGORY_FIELDS = [
    "submission_category",
    "category",
    "categories",
    "subject_area",
    "subject_areas",
    "primary_area",
    "area",
    "areas",
    "track",
    "tracks",
    "paper_type",
    "submission_type",
    "type",
]

ABSTRACT_FIELDS = [
    "abstract",
    "Abstract",
]

TITLE_FIELDS = [
    "title",
    "Title",
]


# ============================================================
# HELPERS
# ============================================================

def extract_forum_id(url_or_id):
    """
    Handles:
      - https://openreview.net/forum?id=ABC123
      - https://openreview.net/forum?id=ABC123&noteId=...
      - raw OpenReview forum/note IDs
    """
    if pd.isna(url_or_id):
        return None

    text = str(url_or_id).strip()

    if not text:
        return None

    match = re.search(r"[?&]id=([^&]+)", text)
    if match:
        return match.group(1)

    return text


def unwrap_openreview_value(value):
    """
    OpenReview API v2 often stores content fields like:
        content["title"]["value"]

    Older structures may store:
        content["title"] = "Some title"

    This function handles both.
    """
    if value is None:
        return None

    if isinstance(value, dict):
        if "value" in value:
            return value["value"]

        # Sometimes fields may be nested differently.
        # Fall back to a readable representation.
        return value

    return value


def content_value(content, key):
    if not content or key not in content:
        return None

    return unwrap_openreview_value(content[key])


def first_existing_field(content, candidates):
    for key in candidates:
        value = content_value(content, key)

        if value not in [None, "", []]:
            return value

    return None


def clean_for_excel(text):
    """
    Remove characters that Excel/openpyxl cannot write.
    Also trims very long cell values to Excel's 32,767 character limit.
    """
    if text is None:
        return ""

    text = str(text)

    # Remove illegal Excel control characters
    text = ILLEGAL_CHARACTERS_RE.sub("", text)

    # Excel cell character limit
    if len(text) > 32767:
        text = text[:32760] + "..."

    return text


def normalize_value(value):
    """
    Converts OpenReview values into spreadsheet-friendly strings.
    """
    if value is None:
        return ""

    value = unwrap_openreview_value(value)

    if value is None:
        return ""

    if isinstance(value, list):
        text = "; ".join(clean_for_excel(x) for x in value)
        return clean_for_excel(text)

    if isinstance(value, dict):
        text = "; ".join(f"{clean_for_excel(k)}: {clean_for_excel(v)}" for k, v in value.items())
        return clean_for_excel(text)

    return clean_for_excel(value)


def print_short(label, value, max_len=160):
    text = normalize_value(value)
    if len(text) > max_len:
        text = text[:max_len] + "..."
    print(f"    {label}: {text}")


def clean_dataframe_for_excel(out_df):
    """
    Apply Excel-safe cleaning to all object/string cells.
    """
    cleaned = out_df.copy()

    for col in cleaned.columns:
        if cleaned[col].dtype == "object":
            cleaned[col] = cleaned[col].map(
                lambda x: clean_for_excel(x) if pd.notna(x) else x
            )

    return cleaned


def save_partial(df, upto_index, keywords_out, categories_out, title_out, abstract_out):
    """
    Save only rows processed so far.
    """
    temp_df = df.iloc[:upto_index].copy()

    temp_df["openreview_keywords"] = keywords_out
    temp_df["submission_category"] = categories_out
    temp_df["openreview_title"] = title_out
    temp_df["openreview_abstract"] = abstract_out

    temp_df = clean_dataframe_for_excel(temp_df)

    temp_df.to_excel(PARTIAL_OUTPUT_FILE, index=False)
    print(f"    Saved partial file: {PARTIAL_OUTPUT_FILE}")


def get_note_with_fallback(forum_id):
    """
    Try API v2 client first.

    If your venue is old and this fails, you may need API v1.
    This script keeps the main path simple, but the error message
    will make it clear which forum failed.
    """
    return client.get_note(forum_id)


# ============================================================
# MAIN
# ============================================================

def main():
    print("=" * 80)
    print("OpenReview sheet enrichment")
    print("=" * 80)
    print(f"Input file:  {INPUT_FILE}")
    print(f"Output file: {OUTPUT_FILE}")
    print()

    if not os.path.exists(INPUT_FILE):
        raise FileNotFoundError(f"Could not find input file: {INPUT_FILE}")

    df = pd.read_excel(INPUT_FILE)

    if FORUM_COLUMN not in df.columns:
        raise ValueError(
            f"Missing required column '{FORUM_COLUMN}'. "
            f"Available columns: {list(df.columns)}"
        )

    total = len(df)

    keywords_out = []
    categories_out = []
    title_out = []
    abstract_out = []

    failed = []

    start_time = time.time()

    for i, (_, row) in enumerate(df.iterrows(), start=1):
        forum_id = extract_forum_id(row.get(FORUM_COLUMN))

        sheet_title = row.get(TITLE_COLUMN, "")
        sheet_number = row.get(NUMBER_COLUMN, "")

        print(f"[{i}/{total}] Processing submission {sheet_number}")
        if sheet_title not in [None, ""]:
            print(f"    sheet title: {sheet_title}")
        print(f"    forum_id: {forum_id}")

        if not forum_id:
            print("    No forum ID found, skipping.")

            keywords_out.append("")
            categories_out.append("")
            title_out.append("")
            abstract_out.append("")

            failed.append({
                "row": i,
                "forum_id": forum_id,
                "reason": "No forum ID found",
            })

            continue

        try:
            print("    contacting OpenReview...")
            note = get_note_with_fallback(forum_id)
            print("    received response from OpenReview")

            content = note.content

            keywords = first_existing_field(content, KEYWORD_FIELDS)
            category = first_existing_field(content, CATEGORY_FIELDS)
            title = first_existing_field(content, TITLE_FIELDS)
            abstract = first_existing_field(content, ABSTRACT_FIELDS)

            keywords_out.append(normalize_value(keywords))
            categories_out.append(normalize_value(category))
            title_out.append(normalize_value(title))
            abstract_out.append(normalize_value(abstract))

            print("    OK")
            print_short("keywords", keywords)
            print_short("category", category)
            print_short("openreview title", title)

            # Useful debugging: show available content fields if category is empty
            if not normalize_value(category):
                available_fields = ", ".join(sorted(content.keys()))
                print(f"    category not found. Available fields: {available_fields}")

        except Exception as e:
            print(f"    FAILED: {e}")

            keywords_out.append("")
            categories_out.append("")
            title_out.append("")
            abstract_out.append("")

            failed.append({
                "row": i,
                "forum_id": forum_id,
                "reason": str(e),
            })

        elapsed = time.time() - start_time
        avg_per_item = elapsed / i
        remaining = total - i
        eta_seconds = avg_per_item * remaining

        print(f"    elapsed: {elapsed / 60:.1f} min | ETA: {eta_seconds / 60:.1f} min")

        if i % SAVE_EVERY == 0:
            save_partial(
                df=df,
                upto_index=i,
                keywords_out=keywords_out,
                categories_out=categories_out,
                title_out=title_out,
                abstract_out=abstract_out,
            )

        print("-" * 80)

    df["openreview_keywords"] = keywords_out
    df["submission_category"] = categories_out
    df["openreview_title"] = title_out
    df["openreview_abstract"] = abstract_out

    df = clean_dataframe_for_excel(df)
    df.to_excel(OUTPUT_FILE, index=False)

    print()
    print("=" * 80)
    print("Done")
    print("=" * 80)
    print(f"Wrote final file: {OUTPUT_FILE}")

    if failed:
        failed_df = pd.DataFrame(failed)
        failed_file = "openreview_failed_rows.xlsx"
        failed_df.to_excel(failed_file, index=False)

        print(f"Failed rows: {len(failed)}")
        print(f"Wrote failure log: {failed_file}")
    else:
        print("No failed rows.")


if __name__ == "__main__":
    main()