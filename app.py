import os
import streamlit as st
import pdfplumber
from pptx import Presentation
from io import BytesIO
import google.generativeai as genai
import json
import tiktoken

# Constants
MAX_TOKENS = 1900000
TOKENS_PER_PAGE_ESTIMATE = 1500
MAX_PAGES_TOTAL = MAX_TOKENS // TOKENS_PER_PAGE_ESTIMATE

# Configure Gemini
GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY", "YOUR_GEMINI_API_KEY")
genai.configure(api_key=GEMINI_API_KEY)

# UI Setup
st.title("AI-Powered PPT Generator with Custom Theme")
st.markdown("""
### Guidelines for PDF Upload:
- Maximum combined length: ~1,250 pages
- For larger books: Upload specific chapters
- Large PDFs (>500 pages): Upload alone
- Multiple smaller PDFs: Combined if under limit
- System shows token count and skips if exceeded
""")

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
        st.info(f"Processing {file.name}...")
        pdf_bytes = file.read()
        text = extract_text_from_pdf(BytesIO(pdf_bytes))
        tokens = count_tokens(text)
        
        if total_tokens + tokens > MAX_TOKENS:
            st.warning(f"Skipping {file.name} - would exceed token limit. Current: {total_tokens:,}")
            continue
            
        combined_text += f"\n\nSource: {file.name}\n{text}"
        total_tokens += tokens
        processed_files.append(file.name)
        st.success(f"Processed {file.name}: {tokens:,} tokens")
    
    st.write(f"Total tokens processed: {total_tokens:,}")
    return combined_text, total_tokens, processed_files

def call_gemini_api_for_slides(source_text, topic, num_slides):
    prompt = f"""
    Create a {num_slides}-slide presentation about "{topic}".
    For each slide provide:
    - Short title (max 8 words)
    - 3-5 bullet points
    Output as JSON only:
    {{
      "slides": [
        {{
          "title": "Title",
          "bullets": ["Point 1", "Point 2", "Point 3"]
        }}
      ]
    }}
    Source text: {source_text}
    """
    
    try:
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
        
        content = response.text.strip()
        if content.startswith("```") and content.endswith("```"):
            content = content.split("```")[1]
            if content.startswith("json\n"):
                content = content[5:]
            content = content.strip()
            
        return json.loads(content)
    except Exception as e:
        st.error(f"API Error: {str(e)}")
        raise

def create_ppt_from_slides(slides_data, template_ppt):
    prs = Presentation(template_ppt)
    slide_layout = prs.slide_layouts[1]  # Title and Content layout
    
    for slide_info in slides_data["slides"]:
        slide = prs.slides.add_slide(slide_layout)
        slide.shapes.title.text = slide_info["title"]
        
        body_shape = slide.shapes.placeholders[1]
        tf = body_shape.text_frame
        tf.text = ""  # Clear default
        
        for bullet in slide_info["bullets"]:
            p = tf.add_paragraph()
            p.text = bullet
            p.level = 0
    
    pptx_stream = BytesIO()
    prs.save(pptx_stream)
    pptx_stream.seek(0)
    return pptx_stream

# UI Inputs
uploaded_files = st.file_uploader("Upload PDF(s)", type=["pdf"], accept_multiple_files=True)
uploaded_template = st.file_uploader("Upload PPT template (optional)", type=["pptx"])
topic = st.text_input("Presentation topic:")
num_slides = st.number_input("Number of slides:", 1, 50, 10)
generate_button = st.button("Generate PPT")

# Main execution
if generate_button:
    if not uploaded_files:
        st.error("Please upload at least one PDF")
    elif not topic:
        st.error("Please enter a topic")
    else:
        try:
            with st.spinner("Processing..."):
                st.info("Step 1: Processing PDFs...")
                source_text, total_tokens, processed_files = process_pdfs(uploaded_files)
                
                if total_tokens == 0:
                    st.error("No content could be processed")
                else:
                    st.info("Step 2: Generating slides content...")
                    slides_data = call_gemini_api_for_slides(source_text, topic, num_slides)
                    
                    st.info("Step 3: Creating PowerPoint...")
                    if uploaded_template:
                        template_bytes = uploaded_template.getvalue()
                        template_ppt = BytesIO(template_bytes)
                    else:
                        try:
                            with open("template.pptx", "rb") as f:
                                template_ppt = BytesIO(f.read())
                        except FileNotFoundError:
                            st.error("Default template.pptx not found in application directory")
                            raise

                    ppt_file = create_ppt_from_slides(slides_data, template_ppt)
                    
                    st.success("✅ Presentation generated!")
                    st.download_button(
                        "⬇️ Download PPT",
                        ppt_file,
                        file_name="presentation.pptx",
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                    )
        except Exception as e:
            st.error(f"Error: {str(e)}")
