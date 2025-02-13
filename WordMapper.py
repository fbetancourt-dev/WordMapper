import zipfile
from lxml import etree
import re
from collections import defaultdict


def extract_tracked_changes_from_docx(file_path, debug=False):
    changes = []
    last_sad_id = None  # Store the most recent SAD ID found
    detected_srs_mappings = defaultdict(set)  # Store detected mappings grouped by SAD
    detected_srs_removals = defaultdict(set)  # Store removed mappings grouped by SAD
    deleted_sad_sections = set()
    existing_sad_sections = set()  # To track SAD IDs that still exist
    unique_changes = set()  # To avoid duplicates

    if debug:
        print(f"Opening document: {file_path}")

    # Open the .docx file as a zip archive
    with zipfile.ZipFile(file_path) as docx_zip:
        # Read the document.xml file
        with docx_zip.open("word/document.xml") as document_xml:
            tree = etree.parse(document_xml)

            # Namespaces used in Word documents
            namespaces = {
                "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            }

            # Detect paragraphs
            paragraphs = tree.xpath("//w:p", namespaces=namespaces)

            for para in paragraphs:
                para_text = "".join(
                    para.xpath(".//w:t/text()", namespaces=namespaces)
                ).strip()
                if debug:
                    print(f"Processing paragraph: {para_text}")

                # Check if the paragraph contains an SAD ID and update the last found SAD ID
                sad_matches = re.findall(r"SAD-\d+", para_text)
                if sad_matches:
                    for sad in sad_matches:
                        existing_sad_sections.add(sad)  # Track existing SADs
                    last_sad_id = sad_matches[-1]  # Always retain the last SAD ID found
                    if debug:
                        print(f"Updated last SAD ID: {last_sad_id}")

                # Detect insertions and deletions
                insertions = para.xpath(".//w:ins", namespaces=namespaces)
                deletions = para.xpath(".//w:del", namespaces=namespaces)
                covers_deleted = any(
                    "Covers:"
                    in "".join(
                        dele.xpath(
                            ".//w:t/text() | .//w:delText/text()", namespaces=namespaces
                        )
                    )
                    for dele in deletions
                )

                inserted_srs = set()
                deleted_srs = set()

                # Capture deleted SAD sections **only if the SAD does NOT exist elsewhere**
                if deletions and sad_matches:
                    for sad in sad_matches:
                        if (
                            sad not in existing_sad_sections
                            and sad not in deleted_sad_sections
                        ):
                            deleted_sad_sections.add(sad)
                            changes.append(f"Deleted {sad}")
                            if debug:
                                print(f"Deleted SAD Section: {sad}")

                # Process insertions inside Covers section
                for ins in insertions:
                    ins_text = "".join(
                        ins.xpath(".//w:t/text()", namespaces=namespaces)
                    ).strip()
                    srs_matches = re.findall(r"SRS-\d+", ins_text)
                    for srs in srs_matches:
                        inserted_srs.add(srs)
                        sad_to_map = last_sad_id if last_sad_id else "Unknown SAD"
                        detected_srs_mappings[sad_to_map].add(srs)
                        if debug:
                            print(f"Inserted in Covers: {srs} mapped to {sad_to_map}")

                # Process deletions inside Covers section only if Covers was not fully removed
                if not covers_deleted:
                    for dele in deletions:
                        del_text = "".join(
                            dele.xpath(
                                ".//w:t/text() | .//w:delText/text()",
                                namespaces=namespaces,
                            )
                        ).strip()
                        srs_matches = re.findall(r"SRS-\d+", del_text)
                        for srs in srs_matches:
                            deleted_srs.add(srs)
                            sad_to_map = last_sad_id if last_sad_id else "Unknown SAD"
                            if (
                                sad_to_map in existing_sad_sections
                            ):  # Ensure the SAD still exists
                                detected_srs_removals[sad_to_map].add(srs)
                                if debug:
                                    print(
                                        f"Deleted in Covers: {srs} removed from {sad_to_map}"
                                    )

    # Format grouped mappings for output
    for sad_id, srs_set in detected_srs_mappings.items():
        grouped_srs = ", ".join(sorted(srs_set))
        changes.append(f"{grouped_srs} mapped to {sad_id}")

    # Format grouped removals for output
    for sad_id, srs_set in detected_srs_removals.items():
        grouped_srs = ", ".join(sorted(srs_set))
        changes.append(f"{grouped_srs} removed from {sad_id}")

    return changes


def main():
    file_path = "TE1605_R000570_SHS_SAD.docx"  # Ensure this is the correct file path
    debug_mode = True  # Set to True for debugging output
    changes = extract_tracked_changes_from_docx(file_path, debug=debug_mode)

    # Display the changes, ensuring correct formatting
    formatted_changes = []
    for change in changes:
        formatted_changes.append(change)

    for change in formatted_changes:
        print(change)


if __name__ == "__main__":
    main()
