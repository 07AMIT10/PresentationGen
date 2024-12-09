import os
import streamlit as st
import pdfplumber
from pptx import Presentation
from pptx.util import Pt
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
        try:
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
        except Exception as e:
            st.error(f"Error processing {file.name}: {str(e)}")
    
    st.write(f"Total tokens processed: {total_tokens:,}")
    return combined_text, total_tokens, processed_files


def validate_ppt(uploaded_template):
    """
    Validates if the uploaded PowerPoint file is legitimate.
    """
    try:
        Presentation(uploaded_template)
        return True
    except Exception:
        return False


def extract_template_info(prs):
    """
    Extract layout and placeholder information from the PowerPoint template.
    """
    layout_info = []
    for layout in prs.slide_layouts:
        placeholders = []
        for placeholder in layout.placeholders:
            placeholders.append({
                "name": placeholder.name,
                "type": placeholder.placeholder_format.type,
                "idx": placeholder.placeholder_format.idx,
            })
        layout_info.append({"layout": layout, "placeholders": placeholders})
    return layout_info


def apply_template_style(slide, slide_info, content):
    """
    Apply template styles (e.g., font, placeholder usage) to generated content.
    """
    if "title" in content and slide.shapes.title:
        slide.shapes.title.text = content["title"]
    
    for placeholder in slide.placeholders:
        if placeholder.placeholder_format.type == 1:  # Body placeholder
            text_frame = placeholder.text_frame
            text_frame.text = ""  # Clear default text
            for bullet in content.get("bullets", []):
                p = text_frame.add_paragraph()
                p.text = bullet
                p.font.size = Pt(14)  # Adjust font size to match the template
            break


def create_ppt_from_template(slides_data, template_ppt, num_slides_required):
    """
    Generate a PowerPoint presentation based on the provided template.
    Adjusts the number of slides to match the required count.
    """
    prs = Presentation(template_ppt)
    layout_info = extract_template_info(prs)
    
    # Remove redundant slides
    while len(prs.slides) > num_slides_required:
        slide_to_remove = prs.slides[len(prs.slides) - 1]
        prs.slides._sldIdLst.remove(slide_to_remove._element)

    # Duplicate slides if more are needed
    while len(prs.slides) < num_slides_required:
        prs.slides.add_slide(prs.slides[0].slide_layout)

    for idx, slide_info in enumerate(slides_data["slides"]):
        if idx < len(prs.slides):
            slide = prs.slides[idx]
        else:
            layout = next(
                (li["layout"] for li in layout_info if len(li["placeholders"]) > 1),
                prs.slide_layouts[0]
            )
            slide = prs.slides.add_slide(layout)
        apply_template_style(slide, layout_info, slide_info)
    
    pptx_stream = BytesIO()
    prs.save(pptx_stream)
    pptx_stream.seek(0)
    return pptx_stream


def call_gemini_api_for_slides(source_text, topic, num_slides):
    """
    Generates slides content using the Gemini API.
    """
    prompt = f"""
    Create a {num_slides}-slide presentation about "{topic}".
    For each slide provide:
    - Short title (max 8 words)
    - 3-5 bullet points
    Output as JSON only:
    {{
      "slides": [
        {{"title": "Title", "bullets": ["Point 1", "Point 2", "Point 3"]}}
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
    elif uploaded_template and not validate_ppt(uploaded_template):
        st.error("Uploaded template is not a valid PowerPoint file.")
    else:
        try:
            with st.spinner("Processing PDFs..."):
                source_text, total_tokens, processed_files = process_pdfs(uploaded_files)
                
                if total_tokens == 0:
                    st.error("No content could be processed.")
                else:
                    slides_data = call_gemini_api_for_slides(source_text, topic, num_slides)
                    if uploaded_template:
                        template_bytes = uploaded_template.getvalue()
                        template_ppt = BytesIO(template_bytes)
                    else:
                        st.error("No valid template uploaded.")
                        raise Exception("No template available.")

                    ppt_file = create_ppt_from_template(slides_data, template_ppt, num_slides)
                    st.success("✅ Presentation generated!")
                    st.download_button(
                        "⬇️ Download PPT",
                        ppt_file,
                        file_name="presentation.pptx",
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                    )
        except Exception as e:
            st.error(f"An error occurred: {str(e)}")
