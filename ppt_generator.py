"""
AI PPT Generator - Main Logic
This file contains all the functions needed to generate PowerPoint presentations using OpenAI.

Flow:
1. User provides a topic/prompt
2. OpenAI generates structured slide content
3. We parse the content and create PPT slides
4. Blank rectangles are added where images should go
"""

import os
from openai import OpenAI
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from dotenv import load_dotenv
import json

# ========== STEP 1: LOAD ENVIRONMENT VARIABLES ==========
# This loads your OpenAI API key from the .env file
load_dotenv()
api_key = os.getenv('OPENAI_API_KEY')
os.environ['OPENAI_API_KEY'] = api_key
client = OpenAI()


def generate_slide_content(client, topic, num_slides=5):
    """
    Uses OpenAI to generate structured content for PowerPoint slides.

    Args:
        client: OpenAI client object
        topic: The topic/prompt provided by the user
        num_slides: Number of slides to generate (default: 5)

    Returns:
        List of dictionaries, each containing slide information:
        - title: Slide title
        - content: List of bullet points
        - notes: Speaker notes
        - has_image: Whether this slide should have an image placeholder
    """

    # Create a detailed prompt for OpenAI to generate structured content
    prompt = f"""
Create content for a {num_slides}-slide PowerPoint presentation on the topic: "{topic}"

For each slide, provide:
1. A clear, engaging title
2. 3-5 concise bullet points (keep each point to 1-2 lines)
3. Brief speaker notes (2-3 sentences)
4. Whether the slide should have an image (yes/no)

Format your response as a JSON array where each slide is an object with these keys:
- "title": string
- "content": array of strings (bullet points)
- "notes": string
- "has_image": boolean

Make the content professional, informative, and well-structured.
The first slide should be a title slide with just the main title and subtitle.
"""

    # Call OpenAI API to generate the content
    print("Calling OpenAI API to generate slide content...")
    response = client.chat.completions.create(
        model="gpt-4o-2024-08-06",
        messages=[
            {"role": "system", "content": "You are a professional presentation designer And A Teacher with 6 yr Experience who creates engaging, well-structured slide content in detail for teaching class 1-12."},
            {"role": "user", "content": prompt}
        ],
        temperature=0.3,
        max_tokens=4000
    )

    # Extract the generated text from the API response
    generated_text = response.choices[0].message.content

    # Parse the JSON response from OpenAI
    # Sometimes the response has markdown code blocks, so we clean it first
    if "```json" in generated_text:
        # Extract JSON from markdown code block
        generated_text = generated_text.split("```json")[1].split("```")[0].strip()
    elif "```" in generated_text:
        generated_text = generated_text.split("```")[1].split("```")[0].strip()

    try:
        slides_data = json.loads(generated_text)
        print(f"Successfully generated content for {len(slides_data)} slides!")
        return slides_data
    except json.JSONDecodeError as e:
        print(f"Error parsing OpenAI response: {e}")
        print("Raw response:", generated_text)
        raise


def create_presentation_from_template(template_name, slides_data, output_filename):
    """
    Creates a PowerPoint presentation using a template and the generated content.

    Args:
        template_name: Name of the template to use ("modern_blue" or "elegant_dark")
        slides_data: List of slide data dictionaries from generate_slide_content()
        output_filename: Name for the output PPT file (e.g., "my_presentation.pptx")

    Returns:
        Path to the created presentation file
    """

    # Load the selected template
    template_path = f"templates/{template_name}.pptx"

    if not os.path.exists(template_path):
        raise FileNotFoundError(f"Template not found: {template_path}")

    print(f"Loading template: {template_name}")
    prs = Presentation(template_path)

    # Define color schemes for each template
    if template_name == "modern_blue":
        text_color = RGBColor(26, 35, 126)      # Dark blue
        accent_color = RGBColor(33, 150, 243)   # Bright blue
        bg_color = RGBColor(255, 255, 255)      # White
    elif template_name == "elegant_dark":
        text_color = RGBColor(250, 250, 250)    # Off-white
        accent_color = RGBColor(255, 193, 7)    # Gold
        bg_color = RGBColor(18, 18, 18)         # Almost black
    else:
        # Default colors if template name doesn't match
        text_color = RGBColor(0, 0, 0)
        accent_color = RGBColor(0, 120, 215)
        bg_color = RGBColor(255, 255, 255)

    # Clear existing slides from template (we'll create new ones)
    # Keep the layouts but remove the sample slides
    while len(prs.slides) > 0:
        rId = prs.slides._sldIdLst[0].rId
        prs.part.drop_rel(rId)
        del prs.slides._sldIdLst[0]

    # Create slides based on the generated content
    for i, slide_data in enumerate(slides_data):
        print(f"Creating slide {i+1}: {slide_data['title']}")

        # Add a blank slide
        slide = prs.slides.add_slide(prs.slide_layouts[6])  # 6 = blank layout

        # Determine if this is the title slide (first slide)
        is_title_slide = (i == 0)

        if is_title_slide:
            # ===== TITLE SLIDE LAYOUT =====
            # Add background elements based on template
            if template_name == "modern_blue":
                # Blue header bar
                top_bar = slide.shapes.add_shape(1, Inches(0), Inches(0), Inches(10), Inches(2))
                top_bar.fill.solid()
                top_bar.fill.fore_color.rgb = RGBColor(26, 35, 126)
                top_bar.line.fill.background()

                # Light background
                bg = slide.shapes.add_shape(1, Inches(0), Inches(2), Inches(10), Inches(5.5))
                bg.fill.solid()
                bg.fill.fore_color.rgb = RGBColor(245, 245, 245)
                bg.line.fill.background()

                # Accent line
                accent = slide.shapes.add_shape(1, Inches(0), Inches(2), Inches(10), Inches(0.1))
                accent.fill.solid()
                accent.fill.fore_color.rgb = accent_color
                accent.line.fill.background()

                title_color = RGBColor(255, 255, 255)
                subtitle_color = text_color

            elif template_name == "elegant_dark":
                # Dark background
                bg = slide.shapes.add_shape(1, Inches(0), Inches(0), Inches(10), Inches(7.5))
                bg.fill.solid()
                bg.fill.fore_color.rgb = bg_color
                bg.line.fill.background()

                # Gold accent strip
                accent_strip = slide.shapes.add_shape(1, Inches(8), Inches(0), Inches(2), Inches(7.5))
                accent_strip.fill.solid()
                accent_strip.fill.fore_color.rgb = accent_color
                accent_strip.line.fill.background()
                accent_strip.rotation = 15

                title_color = text_color
                subtitle_color = text_color

            # Add title text box
            title_box = slide.shapes.add_textbox(Inches(1), Inches(2.5), Inches(8), Inches(1.5))
            title_frame = title_box.text_frame
            title_frame.text = slide_data['title']
            title_para = title_frame.paragraphs[0]
            title_para.font.size = Pt(44)
            title_para.font.bold = True
            title_para.font.color.rgb = title_color
            title_para.alignment = PP_ALIGN.CENTER

            # Add subtitle if present in content
            if slide_data.get('content') and len(slide_data['content']) > 0:
                subtitle_box = slide.shapes.add_textbox(Inches(1), Inches(4.2), Inches(8), Inches(1))
                subtitle_frame = subtitle_box.text_frame
                subtitle_frame.text = slide_data['content'][0]
                subtitle_para = subtitle_frame.paragraphs[0]
                subtitle_para.font.size = Pt(24)
                subtitle_para.font.color.rgb = subtitle_color
                subtitle_para.alignment = PP_ALIGN.CENTER

        else:
            # ===== CONTENT SLIDE LAYOUT =====
            # Add background elements
            if template_name == "modern_blue":
                # Header bar
                header = slide.shapes.add_shape(1, Inches(0), Inches(0), Inches(10), Inches(1))
                header.fill.solid()
                header.fill.fore_color.rgb = RGBColor(26, 35, 126)
                header.line.fill.background()

                # White content area
                content_bg = slide.shapes.add_shape(1, Inches(0), Inches(1), Inches(10), Inches(6.5))
                content_bg.fill.solid()
                content_bg.fill.fore_color.rgb = RGBColor(255, 255, 255)
                content_bg.line.fill.background()

                # Blue sidebar
                sidebar = slide.shapes.add_shape(1, Inches(0), Inches(1), Inches(0.2), Inches(6.5))
                sidebar.fill.solid()
                sidebar.fill.fore_color.rgb = accent_color
                sidebar.line.fill.background()

                header_text_color = RGBColor(255, 255, 255)
                content_text_color = text_color

            elif template_name == "elegant_dark":
                # Dark background
                bg = slide.shapes.add_shape(1, Inches(0), Inches(0), Inches(10), Inches(7.5))
                bg.fill.solid()
                bg.fill.fore_color.rgb = bg_color
                bg.line.fill.background()

                # Dark gray header
                header_section = slide.shapes.add_shape(1, Inches(0), Inches(0), Inches(10), Inches(1.2))
                header_section.fill.solid()
                header_section.fill.fore_color.rgb = RGBColor(33, 33, 33)
                header_section.line.fill.background()

                # Gold accent line
                gold_line = slide.shapes.add_shape(1, Inches(0.5), Inches(1.1), Inches(2), Inches(0.08))
                gold_line.fill.solid()
                gold_line.fill.fore_color.rgb = accent_color
                gold_line.line.fill.background()

                header_text_color = text_color
                content_text_color = text_color

            # Add slide title
            title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.25), Inches(9), Inches(0.7))
            title_frame = title_box.text_frame
            title_frame.text = slide_data['title']
            title_para = title_frame.paragraphs[0]
            title_para.font.size = Pt(32)
            title_para.font.bold = True
            title_para.font.color.rgb = header_text_color

            # Decide layout: if image needed, split content and image area
            has_image = slide_data.get('has_image', False)

            if has_image:
                # Split slide: left side for content, right side for image
                content_left = Inches(0.5)
                content_width = Inches(4.5)
                image_left = Inches(5.5)
                image_width = Inches(4)
                image_top = Inches(2)
                image_height = Inches(4)
            else:
                # Full width for content
                content_left = Inches(0.7)
                content_width = Inches(8.6)

            # Add bullet points
            content_top = Inches(1.8)
            content_height = Inches(4.5)

            content_box = slide.shapes.add_textbox(content_left, content_top, content_width, content_height)
            text_frame = content_box.text_frame
            text_frame.word_wrap = True

            # Add each bullet point
            for j, point in enumerate(slide_data.get('content', [])):
                if j == 0:
                    # First paragraph already exists
                    para = text_frame.paragraphs[0]
                else:
                    # Add new paragraph for subsequent points
                    para = text_frame.add_paragraph()

                para.text = point
                para.level = 0  # Bullet level (0 = main bullet)
                para.font.size = Pt(18)
                para.font.color.rgb = content_text_color
                para.space_before = Pt(12)

            # Add image placeholder if needed
            if has_image:
                # Create a blank rectangle as image placeholder
                image_placeholder = slide.shapes.add_shape(
                    1,  # Rectangle shape
                    image_left, image_top,
                    image_width, image_height
                )
                # Style the placeholder
                image_placeholder.fill.solid()
                image_placeholder.fill.fore_color.rgb = RGBColor(220, 220, 220)  # Light gray
                image_placeholder.line.color.rgb = RGBColor(150, 150, 150)
                image_placeholder.line.width = Pt(2)

                # Add "Image Placeholder" text
                if image_placeholder.has_text_frame:
                    frame = image_placeholder.text_frame
                    frame.text = "[Image Placeholder]"
                    p = frame.paragraphs[0]
                    p.font.size = Pt(16)
                    p.font.color.rgb = RGBColor(100, 100, 100)
                    p.alignment = PP_ALIGN.CENTER

        # Add speaker notes to the slide
        notes_slide = slide.notes_slide
        notes_frame = notes_slide.notes_text_frame
        notes_frame.text = slide_data.get('notes', '')

    # Create output directory if it doesn't exist
    os.makedirs('output', exist_ok=True)

    # Save the presentation
    output_path = f"output/{output_filename}"
    prs.save(output_path)

    print(f"\nPresentation saved successfully: {output_path}")
    return output_path


def generate_ppt(topic, template_name="modern_blue", num_slides=5, output_filename=None):
    """
    Main function that orchestrates the entire PPT generation process.
    This is the function you'll call from the Streamlit app.

    Args:
        topic: The topic/prompt for the presentation
        template_name: Which template to use ("modern_blue" or "elegant_dark")
        num_slides: Number of slides to generate
        output_filename: Name for output file (auto-generated if None)

    Returns:
        Path to the generated PowerPoint file
    """

    # Auto-generate filename if not provided
    if output_filename is None:
        # Create a safe filename from the topic
        safe_topic = "".join(c if c.isalnum() else "_" for c in topic)
        safe_topic = safe_topic[:50]  # Limit length
        output_filename = f"{safe_topic}_presentation.pptx"

    # Step 1: Initialize OpenAI client
    print("Step 1: Initializing OpenAI client...")
    client = initialize_openai_client()

    # Step 2: Generate slide content using AI
    print(f"\nStep 2: Generating content for {num_slides} slides on topic: '{topic}'")
    slides_data = generate_slide_content(client, topic, num_slides)

    # Step 3: Create the PowerPoint presentation
    print(f"\nStep 3: Creating PowerPoint with template '{template_name}'")
    output_path = create_presentation_from_template(template_name, slides_data, output_filename)

    print("\n[SUCCESS] AI-powered presentation generated successfully!")
    return output_path


# ========== TEST FUNCTION (optional) ==========
# You can run this file directly to test the generation
if __name__ == "__main__":
    print("AI PPT Generator - Test Mode")
    print("=" * 50)

    # Test parameters
    test_topic = "Artificial Intelligence in Education"
    test_template = "modern_blue"  # or "elegant_dark"
    test_slides = 5

    try:
        result_path = generate_ppt(
            topic=test_topic,
            template_name=test_template,
            num_slides=test_slides
        )
        print(f"\nTest successful! Check your presentation at: {result_path}")
    except Exception as e:
        print(f"\nError during generation: {e}")
        import traceback
        traceback.print_exc()
