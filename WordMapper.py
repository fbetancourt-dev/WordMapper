import zipfile
from lxml import etree
import re


def extract_tracked_changes_from_docx(file_path, debug=False):
    changes = []
    last_sad_id = None  # Store the most recent SAD ID found

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
                    last_sad_id = sad_matches[-1]  # Always retain the last SAD ID found
                    if debug:
                        print(f"Updated last SAD ID: {last_sad_id}")

                # Detect insertions and deletions
                insertions = para.xpath(".//w:ins", namespaces=namespaces)
                deletions = para.xpath(".//w:del", namespaces=namespaces)

                inserted_srs = set()
                deleted_srs = set()

                # Process insertions recursively inside Covers section
                for ins in insertions:
                    ins_text = "".join(
                        ins.xpath(".//w:t/text()", namespaces=namespaces)
                    ).strip()
                    srs_matches = re.findall(r"SRS-\d+", ins_text)
                    for srs in srs_matches:
                        inserted_srs.add(srs)
                        sad_to_map = last_sad_id if last_sad_id else "Unknown SAD"
                        changes.append(f"{srs} mapped to {sad_to_map}")
                        if debug:
                            print(f"Inserted in Covers: {srs} mapped to {sad_to_map}")

                # Process deletions recursively inside Covers section, including nested <w:delText>
                for dele in deletions:
                    del_text = "".join(
                        dele.xpath(
                            ".//w:t/text() | .//w:delText/text()", namespaces=namespaces
                        )
                    ).strip()
                    srs_matches = re.findall(r"SRS-\d+", del_text)
                    for srs in srs_matches:
                        deleted_srs.add(srs)
                        sad_to_map = last_sad_id if last_sad_id else "Unknown SAD"
                        changes.append(f"{srs} removed from {sad_to_map}")
                        if debug:
                            print(f"Deleted in Covers: {srs} removed from {sad_to_map}")

                # Ensure nested changes inside Covers section are correctly extracted
                if "Covers:" in para_text:
                    all_srs_matches = re.findall(r"SRS-\d+", para_text)
                    for srs in all_srs_matches:
                        sad_to_map = last_sad_id if last_sad_id else "Unknown SAD"
                        if srs in inserted_srs:
                            changes.append(f"{srs} mapped to {sad_to_map}")
                            if debug:
                                print(
                                    f"Covers section (Inserted): {srs} mapped to {sad_to_map}"
                                )
                        elif srs in deleted_srs:
                            changes.append(f"{srs} removed from {sad_to_map}")
                            if debug:
                                print(
                                    f"Covers section (Deleted): {srs} removed from {sad_to_map}"
                                )
                        else:
                            changes.append(
                                f"{srs} mapped to {sad_to_map}"
                            )  # Catch missed insertions
                            if debug:
                                print(
                                    f"Covers section (Extra Inserted Check): {srs} mapped to {sad_to_map}"
                                )

    return changes


def main():
    file_path = "TE1605_R000570_SHS_SAD.docx"  # Ensure this is the correct file path
    debug_mode = True  # Set to True for debugging output
    changes = extract_tracked_changes_from_docx(file_path, debug=debug_mode)

    # Display the changes
    for change in changes:
        print(change)


if __name__ == "__main__":
    main()
