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

# Function to interpret similarity scores
def interpret_similarity(score):
    if score == 1.0:
        return "Exact Match! The documents are identical."
    elif 0.8 <= score < 1.0:
        return "High Similarity: The documents are very similar."
    elif 0.5 <= score < 0.8:
        return "Moderate Similarity: The documents share some content."
    elif 0.2 <= score < 0.5:
        return "Low Similarity: The documents have limited overlap."
    else:
        return "Very Low Similarity: The documents are mostly different."

# Function to handle the similarity calculation
def similarity(doc1, doc2, file1=None, file2=None):
    text1 = ""
    text2 = ""

    # Check for file uploads
    if file1 is not None and file1.name.endswith('.docx'):
        text1 = extract_text_from_docx(file1.name)
    elif doc1:
        text1 = doc1
    else:
        return "Please provide either a DOCX file or paste the text for Document 1."

    if file2 is not None and file2.name.endswith('.docx'):
        text2 = extract_text_from_docx(file2.name)
    elif doc2:
        text2 = doc2
    else:
        return "Please provide either a DOCX file or paste the text for Document 2."

    score = calculate_similarity(text1, text2)
    return f"Similarity Score: {score:.2f}\n{interpret_similarity(score)}"

# Create a Gradio interface
with gr.Blocks() as demo:
    gr.Markdown("## 📄 Document Similarity Checker")
    gr.Markdown(
        "Compare two documents by uploading DOCX files or pasting text. The app calculates similarity using MinHash and provides an interpretative score.")
    with gr.Row():
        with gr.Column():
            gr.Markdown("### Document 1")
            file1 = gr.File(label="Upload DOCX File")
            doc1 = gr.Textbox(label="Or Paste Text Here", lines=10, placeholder="Paste document text...")
        with gr.Column():
            gr.Markdown("### Document 2")
            file2 = gr.File(label="Upload DOCX File")
            doc2 = gr.Textbox(label="Or Paste Text Here", lines=10, placeholder="Paste document text...")
    output = gr.Textbox(label="Result", lines=3)
    submit = gr.Button("Check Similarity", variant="primary")

    submit.click(fn=similarity, inputs=[doc1, doc2, file1, file2], outputs=output)

# Launch the Gradio app
demo.launch()
