import os
import streamlit as st
import PyPDF2
import pdfplumber
from pptx import Presentation
from io import BytesIO
import google.generativeai as genai
import json
import tiktoken  # Add to requirements.txt

# Set your Gemini API Key.
# If not set in environment variables, replace "YOUR_GEMINI_API_KEY" with actual key.
GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY", "YOUR_GEMINI_API_KEY")
genai.configure(api_key=GEMINI_API_KEY)

# Add these constants at the top
MAX_TOKENS = 1900000  # Leave buffer from 2M limit
TOKENS_PER_PAGE_ESTIMATE = 1500  # Average estimate
MAX_PAGES_TOTAL = MAX_TOKENS // TOKENS_PER_PAGE_ESTIMATE

# Update Streamlit UI
st.title("AI-Powered PPT Generator with Custom Theme")

st.markdown("""
### Guidelines for PDF Upload:
- Maximum combined length: ~1,250 pages
- For larger books: Upload specific chapters
- Large PDFs (>500 pages): Upload alone
- Multiple smaller PDFs: Combined if under limit
- System shows token count and skips if exceeded
""")

# Multiple file uploader
uploaded_files = st.file_uploader("Upload PDF(s) (Max ~1250 pages total)", 
                                type=["pdf"], 
                                accept_multiple_files=True)

uploaded_template = st.file_uploader("Upload a PPTX template (optional)", type=["pptx"])
topic = st.text_input("Enter the topic for the slides:")
num_slides = st.number_input("Number of slides:", min_value=1, max_value=50, value=10)
generate_button = st.button("Generate PPT")

def extract_text_from_pdf(pdf_bytes):
    all_text = ""
    with pdfplumber.open(pdf_bytes) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if text:
                all_text += text + "\n"
    return all_text

def count_tokens(text):
    encoding = tiktoken.get_encoding("cl100k_base")
    return len(encoding.encode(text))

def process_pdfs(uploaded_files):
    combined_text = ""
    total_tokens = 0
    processed_files = []
    
    for file in uploaded_files:
        pdf_bytes = file.read()
        text = extract_text_from_pdf(BytesIO(pdf_bytes))
        tokens = count_tokens(text)
        
        # Check if adding this file would exceed token limit
        if total_tokens + tokens > MAX_TOKENS:
            st.warning(f"Skipping {file.name} as it would exceed the token limit. Current total: {total_tokens:,} tokens")
            continue
            
        combined_text += f"\n\nSource: {file.name}\n{text}"
        total_tokens += tokens
        processed_files.append(file.name)
        
        st.info(f"Processed {file.name}: {tokens:,} tokens")
    
    st.write(f"Total tokens processed: {total_tokens:,}")
    return combined_text, total_tokens, processed_files

def call_gemini_api_for_slides(source_text, topic, num_slides):
    prompt = f"""
    You are an assistant that creates PowerPoint slide outlines from provided source text. 
    The user wants a presentation on the topic "{topic}" with exactly {num_slides} slides.
    For each slide, provide:
    - A short, compelling slide title (no more than 8 words)
    - 3-5 bullet points of text
    
    Make sure the slides flow logically and cover key points from the provided text. 
    Output ONLY in JSON format like:
    {{
      "slides": [
        {{
          "title": "Slide Title 1",
          "bullets": ["Point 1", "Point 2", "Point 3"]
        }},
        ...
      ]
    }}

    Source Text:
    {source_text}
    """

    model = genai.GenerativeModel('gemini-1.5-pro')
    response = model.generate_content(
        prompt,
        generation_config={
            "temperature": 0.2,
            "top_p": 1,
            "top_k": 32,
            "max_output_tokens": 2048,
        }
    )

    # Clean the response text
    content = response.text.strip()
    
    # Remove Markdown code blocks if present
    if content.startswith("```") and content.endswith("```"):
        # Extract content between code blocks
        content = content.split("```")[1]
        # Remove language identifier if present (e.g., 'json')
        if content.startswith("json\n"):
            content = content[5:]
        content = content.strip()

    try:
        slides_data = json.loads(content)
    except json.JSONDecodeError:
        raise ValueError(f"The model did not return valid JSON. Response was: {content}")
    
    return slides_data
