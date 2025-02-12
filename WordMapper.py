import zipfile
from lxml import etree
import re


def extract_tracked_changes_from_docx(file_path, debug=False):
    changes = []
    sad_ids = []  # Store all SAD IDs encountered

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

                # Check if the paragraph contains an SAD ID and store it
                sad_matches = re.findall(r"SAD-\d+", para_text)
                if sad_matches:
                    sad_ids.extend(sad_matches)
                    if debug:
                        print(f"Found SAD IDs: {sad_matches}")

                # Detect insertions (w:ins) and deletions (w:del)
                insertions = para.xpath(".//w:ins//w:t", namespaces=namespaces)
                deletions = para.xpath(".//w:del", namespaces=namespaces)

                # Process insertions
                for ins in insertions:
                    added_text = ins.text.strip()
                    if added_text.startswith("SRS-"):
                        closest_sad = sad_ids[-1] if sad_ids else "Unknown SAD"
                        changes.append(f"{added_text} mapped to {closest_sad}")
                        if debug:
                            print(f"Inserted: {added_text} mapped to {closest_sad}")

                # Process deletions
                for dele in deletions:
                    deleted_texts = dele.xpath(".//w:t", namespaces=namespaces)
                    for deleted_text in deleted_texts:
                        removed_text = deleted_text.text.strip()
                        if removed_text.startswith("SRS-"):
                            closest_sad = sad_ids[-1] if sad_ids else "Unknown SAD"
                            changes.append(f"{removed_text} removed from {closest_sad}")
                            if debug:
                                print(
                                    f"Deleted: {removed_text} removed from {closest_sad}"
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
