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

INPUT_FILE = "ARR.xlsx"
OUTPUT_FILE = "ARR_enriched.xlsx"
PARTIAL_OUTPUT_FILE = "ARR_enriched_partial.xlsx"

#VENUE_ID = "aclweb.org/ACL/2026/SRW_Direct_Submission"
VENUE_ID = "aclweb.org/ACL/2026/SRW_ARR_Commitment"
PREFERRED_EMAILS_INVITATION_ID = VENUE_ID + "/-/Preferred_Emails"

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


client = openreview.api.OpenReviewClient(
    baseurl="https://api2.openreview.net",
    username=OPENREVIEW_USERNAME,
    password=OPENREVIEW_PASSWORD,
)

print(f"Impersonating venue: {VENUE_ID}")
client.impersonate(VENUE_ID)


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

AUTHOR_NAME_FIELDS = [
    "authors",
    "author_names",
    "names",
]

AUTHOR_ID_FIELDS = [
    "authorids",
    "author_ids",
]


# ============================================================
# HELPERS
# ============================================================

profile_cache = {}


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
    text = ILLEGAL_CHARACTERS_RE.sub("", text)

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
        text = "; ".join(clean_for_excel(x) for x in value if x not in [None, ""])
        return clean_for_excel(text)

    if isinstance(value, dict):
        text = "; ".join(
            f"{clean_for_excel(k)}: {clean_for_excel(v)}"
            for k, v in value.items()
        )
        return clean_for_excel(text)

    return clean_for_excel(value)


def to_list(value):
    """
    Convert OpenReview field values to a clean Python list.
    """
    value = unwrap_openreview_value(value)

    if value is None:
        return []

    if isinstance(value, list):
        return [str(x).strip() for x in value if str(x).strip()]

    if isinstance(value, str):
        # Usually authorids are already a list, but handle semicolon strings too.
        return [x.strip() for x in value.split(";") if x.strip()]

    return [str(value).strip()]


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


def get_author_profiles(author_ids):
    """
    Given a list of OpenReview author IDs / emails, return profile objects.

    Uses venue impersonation + Preferred_Emails invitation so chairs can retrieve
    full preferred emails when permitted by the venue.
    """
    author_ids = to_list(author_ids)

    if not author_ids:
        return []

    missing_ids = [aid for aid in author_ids if aid not in profile_cache]

    if missing_ids:
        try:
            profiles = openreview.tools.get_profiles(
                client,
                missing_ids,
                with_preferred_emails=PREFERRED_EMAILS_INVITATION_ID,
            )

            # Usually returns a list of Profile objects.
            for profile in profiles:
                if profile:
                    profile_cache[getattr(profile, "id", "")] = profile

            # Ensure every requested author ID has an entry, even if no profile found.
            for aid in missing_ids:
                if aid not in profile_cache:
                    matched = None
                    for profile in profiles:
                        if not profile:
                            continue
                        if getattr(profile, "id", None) == aid:
                            matched = profile
                            break
                    profile_cache[aid] = matched

        except Exception as e:
            print(f"    Could not fetch preferred emails for some author profiles: {e}")
            for aid in missing_ids:
                profile_cache[aid] = None

    return [profile_cache.get(aid) for aid in author_ids]


def get_profile_preferred_email(profile):
    """
    Extract full preferred email from an OpenReview profile.

    This should use profile.get_preferred_email() when venue impersonation and
    with_preferred_emails are working.
    """
    if not profile:
        return ""

    try:
        email = profile.get_preferred_email()
        if email:
            return normalize_value(email)
    except Exception:
        pass

    # Fallbacks for older/different openreview-py versions.
    content = getattr(profile, "content", {}) or {}

    for key in ["preferredEmail", "preferred_email"]:
        if key in content:
            return normalize_value(content[key])

    emails = content.get("emails") or content.get("emailsConfirmed")
    if emails:
        return normalize_value(emails)

    return ""


def get_profile_preferred_name(profile):
    """
    Extract preferred/full name from an OpenReview profile.
    """
    if not profile:
        return ""

    try:
        name = profile.get_preferred_name()
        if name:
            return normalize_value(name)
    except Exception:
        pass

    content = getattr(profile, "content", {}) or {}

    names = content.get("names", [])
    if isinstance(names, list) and names:
        preferred_names = [
            n for n in names
            if isinstance(n, dict) and n.get("preferred")
        ]

        if preferred_names:
            name = preferred_names[0]
        else:
            name = names[0]

        if isinstance(name, dict):
            return normalize_value(
                name.get("fullname")
                or " ".join(
                    x for x in [
                        name.get("first"),
                        name.get("middle"),
                        name.get("last"),
                    ]
                    if x
                )
            )

    return normalize_value(getattr(profile, "id", ""))


def save_partial(
    df,
    upto_index,
    keywords_out,
    categories_out,
    title_out,
    abstract_out,
    author_names_out,
    author_ids_out,
    author_emails_out,
):
    """
    Save only rows processed so far.
    """
    temp_df = df.iloc[:upto_index].copy()

    temp_df["openreview_keywords"] = keywords_out
    temp_df["submission_category"] = categories_out
    temp_df["openreview_title"] = title_out
    temp_df["openreview_abstract"] = abstract_out
    temp_df["openreview_author_names"] = author_names_out
    temp_df["openreview_authorids"] = author_ids_out
    temp_df["openreview_author_emails"] = author_emails_out

    temp_df = clean_dataframe_for_excel(temp_df)

    temp_df.to_excel(PARTIAL_OUTPUT_FILE, index=False)
    print(f"    Saved partial file: {PARTIAL_OUTPUT_FILE}")


def get_note_with_fallback(forum_id):
    """
    Try API v2 client.
    """
    return client.get_note(forum_id)


def append_blank_outputs(
    keywords_out,
    categories_out,
    title_out,
    abstract_out,
    author_names_out,
    author_ids_out,
    author_emails_out,
):
    """
    Keep all output lists aligned with the original spreadsheet rows.
    """
    keywords_out.append("")
    categories_out.append("")
    title_out.append("")
    abstract_out.append("")
    author_names_out.append("")
    author_ids_out.append("")
    author_emails_out.append("")


# ============================================================
# MAIN
# ============================================================

def main():
    print("=" * 80)
    print("OpenReview sheet enrichment")
    print("=" * 80)
    print(f"Input file:  {INPUT_FILE}")
    print(f"Output file: {OUTPUT_FILE}")
    print(f"Venue ID:    {VENUE_ID}")
    print(f"Preferred emails invitation: {PREFERRED_EMAILS_INVITATION_ID}")
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
    author_names_out = []
    author_ids_out = []
    author_emails_out = []

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

            append_blank_outputs(
                keywords_out,
                categories_out,
                title_out,
                abstract_out,
                author_names_out,
                author_ids_out,
                author_emails_out,
            )

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

            author_names = first_existing_field(content, AUTHOR_NAME_FIELDS)
            author_ids = first_existing_field(content, AUTHOR_ID_FIELDS)

            author_id_list = to_list(author_ids)
            profiles = get_author_profiles(author_id_list)

            profile_names = [get_profile_preferred_name(p) for p in profiles]
            profile_emails = [get_profile_preferred_email(p) for p in profiles]

            # Keep submitted author names if available; use profile names as fallback.
            if normalize_value(author_names):
                author_names_final = normalize_value(author_names)
            else:
                author_names_final = normalize_value(profile_names)

            author_ids_final = normalize_value(author_id_list)
            author_emails_final = normalize_value(profile_emails)

            keywords_out.append(normalize_value(keywords))
            categories_out.append(normalize_value(category))
            title_out.append(normalize_value(title))
            abstract_out.append(normalize_value(abstract))
            author_names_out.append(author_names_final)
            author_ids_out.append(author_ids_final)
            author_emails_out.append(author_emails_final)

            print("    OK")
            print_short("keywords", keywords)
            print_short("category", category)
            print_short("openreview title", title)
            print_short("authors", author_names_final)
            print_short("author ids", author_ids_final)
            print_short("author emails", author_emails_final)

            # Useful debugging: show available content fields if category is empty.
            if not normalize_value(category):
                available_fields = ", ".join(sorted(content.keys()))
                print(f"    category not found. Available fields: {available_fields}")

            # Useful debugging: if emails are still masked.
            if "****@" in author_emails_final:
                print(
                    "    WARNING: emails still appear masked. "
                    "Check venue impersonation permissions and Preferred_Emails invitation."
                )

        except Exception as e:
            print(f"    FAILED: {e}")

            append_blank_outputs(
                keywords_out,
                categories_out,
                title_out,
                abstract_out,
                author_names_out,
                author_ids_out,
                author_emails_out,
            )

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
                author_names_out=author_names_out,
                author_ids_out=author_ids_out,
                author_emails_out=author_emails_out,
            )

        print("-" * 80)

    df["openreview_keywords"] = keywords_out
    df["submission_category"] = categories_out
    df["openreview_title"] = title_out
    df["openreview_abstract"] = abstract_out
    df["openreview_author_names"] = author_names_out
    df["openreview_authorids"] = author_ids_out
    df["openreview_author_emails"] = author_emails_out

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
        failed_df = clean_dataframe_for_excel(failed_df)
        failed_df.to_excel(failed_file, index=False)

        print(f"Failed rows: {len(failed)}")
        print(f"Wrote failure log: {failed_file}")
    else:
        print("No failed rows.")


if __name__ == "__main__":
    main()
