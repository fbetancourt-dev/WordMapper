import zipfile
from lxml import etree
import re
from collections import defaultdict


def extract_tracked_changes_from_docx(file_path, debug=False):
    changes = []
    last_sdd_id = None  # Store the most recent SDD ID found
    detected_srs_mappings = defaultdict(set)  # Store detected mappings grouped by SDD
    detected_srs_removals = defaultdict(set)  # Store removed mappings grouped by SDD
    deleted_sdd_sections = set()
    existing_sdd_sections = set()  # To track SDD IDs that still exist
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

                # Check if the paragraph contains an SDD ID and update the last found SDD ID
                sdd_matches = re.findall(r"SDD-\d+", para_text)
                if sdd_matches:
                    for sdd in sdd_matches:
                        existing_sdd_sections.add(sdd)  # Track existing SDDs
                    last_sdd_id = sdd_matches[-1]  # Always retain the last SDD ID found
                    if debug:
                        print(f"Updated last SDD ID: {last_sdd_id}")

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

                # Capture deleted SDD sections **only if the SDD does NOT exist elsewhere**
                if deletions and sdd_matches:
                    for sdd in sdd_matches:
                        if (
                            sdd not in existing_sdd_sections
                            and sdd not in deleted_sdd_sections
                        ):
                            deleted_sdd_sections.add(sdd)
                            changes.append(f"Deleted {sdd}")
                            if debug:
                                print(f"Deleted SDD Section: {sdd}")

                # Process insertions inside Covers section
                for ins in insertions:
                    ins_text = "".join(
                        ins.xpath(".//w:t/text()", namespaces=namespaces)
                    ).strip()
                    srs_matches = re.findall(r"SAD-\d+", ins_text)
                    for srs in srs_matches:
                        inserted_srs.add(srs)
                        sdd_to_map = last_sdd_id if last_sdd_id else "Unknown SDD"
                        detected_srs_mappings[sdd_to_map].add(srs)
                        if debug:
                            print(f"Inserted in Covers: {srs} mapped to {sdd_to_map}")

                # Process deletions inside Covers section only if Covers was not fully removed
                if not covers_deleted:
                    for dele in deletions:
                        del_text = "".join(
                            dele.xpath(
                                ".//w:t/text() | .//w:delText/text()",
                                namespaces=namespaces,
                            )
                        ).strip()
                        srs_matches = re.findall(r"SAD-\d+", del_text)
                        for srs in srs_matches:
                            deleted_srs.add(srs)
                            sdd_to_map = last_sdd_id if last_sdd_id else "Unknown SDD"
                            if (
                                sdd_to_map in existing_sdd_sections
                            ):  # Ensure the SDD still exists
                                detected_srs_removals[sdd_to_map].add(srs)
                                if debug:
                                    print(
                                        f"Deleted in Covers: {srs} removed from {sdd_to_map}"
                                    )

    # Format grouped mappings for output, ensuring all SAD are properly retained
    for sdd_id, srs_set in detected_srs_mappings.items():
        grouped_srs = ", ".join(sorted(srs_set, key=lambda x: int(x.split("-")[1])))
        changes.append(f"{grouped_srs} mapped to {sdd_id}")

    # Format grouped removals for output, ensuring all SAD are properly retained
    for sdd_id, srs_set in detected_srs_removals.items():
        grouped_srs = ", ".join(sorted(srs_set, key=lambda x: int(x.split("-")[1])))
        changes.append(f"{grouped_srs} removed from {sdd_id}")

    return changes


def main():
    file_path = "TE1705_R000570_SHS_SDD.docx"  # Ensure this is the correct file path
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
