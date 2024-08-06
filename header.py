from docx import Document

def is_capitalized_correctly(word, exceptions):
    """
    Check if a word is capitalized correctly, allowing for exceptions.
    """
    return word.istitle() or word in exceptions

def analyze_headings(file_path):
    """
    Analyze headings in a Word document to check capitalization.
    """
    # Load the document
    document = Document(file_path)
    
    # Define exceptions (words that should not be capitalized)
    exceptions = {
        "and", "the", "for", "in", "of", "on", "at", "by", "with", 
        "a", "an", "to", "its", "Catia", "Addin.m", "BuaaDieDesignSystem.m", 
        "DateInputDlg", "FEM", "CATIA", "CAA", "ABAQUS","摘","要"
    }
    
    # List to store results
    results = []

    # Iterate through each paragraph in the document
    for paragraph in document.paragraphs:
        # Check the style of the paragraph
        if paragraph.style.name in {"Heading 1", "Heading 2", "Heading 3"}:
            # Split the text into words
            words = paragraph.text.split()
            # Check each word for capitalization
            for word in words:
                if not is_capitalized_correctly(word, exceptions):
                    results.append((paragraph.style.name, paragraph.text, word))

    return results

def main():
    # Path to your Word document
    file_path = 'thesis.docx'
    
    # Analyze the document
    issues = analyze_headings(file_path)
    
    # Output results
    if issues:
        print("The following headings have capitalization issues:")
        for style, heading, word in issues:
            print(f"Style: {style}, Heading: \"{heading}\", Problematic word: \"{word}\"")
    else:
        print("All headings are correctly capitalized.")

if __name__ == "__main__":
    main()
