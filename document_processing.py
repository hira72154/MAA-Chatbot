from docx import Document
import re

file_path = r"C:\Users\user\OneDrive\Documents\Attachments\INQUISITOR SOCIETY.docx"

doc = Document(file_path)

cleaned_text = []

for paragraph in doc.paragraphs:
    text = paragraph.text.strip()

    # Extra spaces remove
    text = re.sub(r"\s+", " ", text)

    if text:
        cleaned_text.append(text)

with open("inquisitor_data_cleaned.txt", "w", encoding="utf-8") as file:
    file.write("\n".join(cleaned_text))

print("Text cleaning completed successfully!")
# Organize cleaned data into sections

sections = {}

current_section = "General Information"
sections[current_section] = []

with open("inquisitor_data_cleaned.txt", "r", encoding="utf-8") as file:
    for line in file:
        line = line.strip()

        if not line:
            continue

        # Detect section headings
        if line in [
            "About Inquisitors Society",
            "Vision",
            "Mission",
            "Membership Information",
            "Events & Activities",
            "Projects & Programs",
            "AI, CSS, PMS, Quiz & Education Services",
            "FAQs",
            "Rules & Policies",
            "Team Information",
            "Contact Details",
            "Social Media Links"
        ]:
            current_section = line
            sections[current_section] = []
        else:
            sections[current_section].append(line)

# Save organized information
with open("inquisitor_data_organized.txt", "w", encoding="utf-8") as file:
    for section, content in sections.items():
        file.write(f"\n### {section}\n")
        for item in content:
            file.write(item + "\n")

print("Data organization completed successfully!")
