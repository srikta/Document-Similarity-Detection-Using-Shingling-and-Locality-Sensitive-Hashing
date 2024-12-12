import docx  # Importing the required module for DOCX extraction
from datasketch import MinHash, MinHashLSH  # Importing MinHash and LSH from datasketch
import gradio as gr  # Importing Gradio for creating the web interface

# Function to extract text from DOCX files
def extract_text_from_docx(docx_path):
    try:
        doc = docx.Document(docx_path)  # Open the DOCX file
        text = "\n".join([para.text for para in doc.paragraphs])  # Extract the text from paragraphs
        return text
    except Exception as e:
        print(f"Error extracting text from DOCX: {str(e)}")
        return ""

# Function to calculate MinHash-based similarity between two texts
def calculate_similarity(doc1, doc2):
    def text_to_shingles(text, k=5):
        # Split the text into k-grams (shingles) of length k
        shingles = set()
        for i in range(len(text) - k + 1):
            shingles.add(text[i:i + k])
        return shingles

    # Generate shingles for both documents
    shingles1 = text_to_shingles(doc1)
    shingles2 = text_to_shingles(doc2)

    # Compute MinHash signatures
    minhash1 = MinHash(num_perm=128)
    minhash2 = MinHash(num_perm=128)
    
    for shingle in shingles1:
        minhash1.update(shingle.encode('utf8'))
    
    for shingle in shingles2:
        minhash2.update(shingle.encode('utf8'))

    # Compute Jaccard similarity using MinHash
    similarity_score = minhash1.jaccard(minhash2)
    return similarity_score

# Function to handle the file upload and similarity calculation
def similarity(file1, file2):
    if file1.name.endswith('.docx'):
        text1 = extract_text_from_docx(file1.name)
    else:
        return "File type not supported. Please upload a DOCX file."

    if file2.name.endswith('.docx'):
        text2 = extract_text_from_docx(file2.name)
    else:
        return "File type not supported. Please upload a DOCX file."

    return calculate_similarity(text1, text2)

# Create a Gradio interface
with gr.Blocks() as demo:
    gr.Markdown("## Document Similarity Checker for DOCX files")
    file1 = gr.File(label="Upload Document 1 (DOCX)")
    file2 = gr.File(label="Upload Document 2 (DOCX)")
    output = gr.Textbox(label="Similarity Score")
    submit = gr.Button("Submit")
    
    submit.click(fn=similarity, inputs=[file1, file2], outputs=output)

# Launching the Streamlit app
demo.launch()
