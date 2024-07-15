from docx import Document

# Load the document
doc = Document('test_worddoc.docx')

# Define a mapping of placeholders to replacement text
placeholder_to_text = {
    'variableA': '1934852',
    'variableB': 'Replacement Text B',
}

# Function to replace placeholders with text
def replace_placeholder(paragraph, placeholder, replacement_text):
    # Iterate through runs in the paragraph
    for run in paragraph.runs:
        if placeholder in run.text:
            # Replace placeholder with replacement text
            run.text = run.text.replace(placeholder, replacement_text)

# Replace placeholders with text
for para in doc.paragraphs:
    for placeholder, replacement_text in placeholder_to_text.items():
        replace_placeholder(para, placeholder, replacement_text)

# Save the modified document
doc.save('modified_document.docx')
