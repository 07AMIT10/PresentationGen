

import os
import streamlit as st
import PyPDF2
import pdfplumber
from pptx import Presentation
from io import BytesIO
import google.generativeai as genai
import json

# Set your Gemini API Key.
# If not set in environment variables, replace "YOUR_GEMINI_API_KEY" with actual key.
GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY", "YOUR_GEMINI_API_KEY")
genai.configure(api_key=GEMINI_API_KEY)

st.title("AI-Powered PPT Generator with Custom Theme")

uploaded_file = st.file_uploader("Upload a PDF (source content)", type=["pdf"])
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

    response = genai.generate_text(
        model="chat-bison-001",
        prompt=prompt,
        temperature=0.2,
        max_output_tokens=1024
    )

    content = response.generations[0].text.strip()

    try:
        slides_data = json.loads(content)
    except json.JSONDecodeError:
        raise ValueError("The model did not return valid JSON. Response was: " + content)
    return slides_data

def create_ppt_from_slides(slides_data, template_ppt):
    prs = Presentation(template_ppt)

    # Use a generic "Title and Content" layout. 
    # Adjust layout index if needed based on your template structure.
    # Commonly:
    # 0: Title Slide
    # 1: Title and Content
    # ...
    slide_layout = prs.slide_layouts[1]

    for slide_info in slides_data["slides"]:
        slide = prs.slides.add_slide(slide_layout)
        title_placeholder = slide.shapes.title
        body_placeholder = slide.placeholders[1]

        title_placeholder.text = slide_info["title"]
        
        body_tf = body_placeholder.text_frame
        body_tf.text = ""  # Clear any default text
        for bullet in slide_info["bullets"]:
            p = body_tf.add_paragraph()
            p.text = bullet
            p.level = 0  # top-level bullet

    output_stream = BytesIO()
    prs.save(output_stream)
    output_stream.seek(0)
    return output_stream

if generate_button and uploaded_file and topic:
    with st.spinner("Processing..."):
        pdf_bytes = uploaded_file.read()
        source_text = extract_text_from_pdf(BytesIO(pdf_bytes))
        
        # Truncate source if too large. Better: summarize if needed.
        if len(source_text) > 15000:
            source_text = source_text[:15000]

        slides_data = call_gemini_api_for_slides(source_text, topic, num_slides)

        # Determine which template to use:
        if uploaded_template is not None:
            template_bytes = uploaded_template.read()
            template_ppt = BytesIO(template_bytes)
        else:
            # Use default template from file
            with open("template.pptx", "rb") as f:
                template_ppt = BytesIO(f.read())

        ppt_file = create_ppt_from_slides(slides_data, template_ppt)

    st.success("Slides generated successfully!")
    st.download_button(
        label="Download PPT",
        data=ppt_file,
        file_name="generated_presentation.pptx",
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
    )
elif generate_button and not uploaded_file:
    st.error("Please upload a PDF first.")
