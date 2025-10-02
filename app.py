"""
AI-Powered PPT Maker - Streamlit App
This is the main user interface for the AI PPT generator.
Run this with: streamlit run app.py
"""

import streamlit as st
import os
from ppt_generator import generate_ppt
from html_generator import generate_html_ppt

# ========== PAGE CONFIGURATION ==========
# Set up the page title and layout
st.set_page_config(
    page_title="AI PPT Maker",
    page_icon="📊",
    layout="wide"
)

# ========== MAIN HEADER ==========
st.title("📊 AI-Powered Presentation Maker")
st.markdown("Generate professional presentations in **PowerPoint** or **HTML with Tailwind CSS** using AI!")
st.markdown("Choose your format: Traditional PowerPoint (.pptx) or Modern HTML with beautiful Tailwind styling")
st.markdown("---")

# ========== CHECK FOR API KEY ==========
# Check if the OpenAI API key is configured
if not os.getenv('OPENAI_API_KEY'):
    st.error("⚠️ OpenAI API key not found!")
    st.info("""
    Please create a `.env` file in the project directory with your OpenAI API key:

    ```
    OPENAI_API_KEY=your_api_key_here
    ```

    Get your API key from: https://platform.openai.com/api-keys
    """)
    st.stop()

# ========== SIDEBAR - SETTINGS ==========
st.sidebar.header("⚙️ Presentation Settings")

# Output Format Selection
output_format = st.sidebar.radio(
    "Output Format",
    ["PowerPoint (.pptx)", "HTML with Tailwind CSS"],
    help="Choose the format for your presentation"
)

is_html = "HTML" in output_format

# Template selection
if is_html:
    template_option = st.sidebar.selectbox(
        "Choose HTML Style",
        ["vibrant", "modern", "dark"],
        help="Select a Tailwind CSS style for your HTML presentation"
    )

    # Display style info
    if template_option == "vibrant":
        st.sidebar.info("🎨 **Vibrant**: Colorful gradient design with purple, blue, green colors")
    elif template_option == "modern":
        st.sidebar.info("🔵 **Modern**: Clean blue gradient design")
    elif template_option == "dark":
        st.sidebar.info("⚫ **Dark**: Professional dark theme with yellow accents")
else:
    template_option = st.sidebar.selectbox(
        "Choose PowerPoint Template",
        ["ocean_breeze", "sunset_glow", "forest_fresh", "royal_purple", "corporate_slate", "vibrant_education", "modern_blue", "elegant_dark"],
        help="Select a professional template for your presentation"
    )

    # Display template preview info
    template_descriptions = {
        "ocean_breeze": "🌊 **Ocean Breeze**: Blue gradient design - Fresh & Educational (RECOMMENDED)",
        "sunset_glow": "🌅 **Sunset Glow**: Orange-pink gradient - Warm & Engaging",
        "forest_fresh": "🌿 **Forest Fresh**: Green tones - Nature & Growth theme",
        "royal_purple": "👑 **Royal Purple**: Purple gradient - Creative & Premium",
        "corporate_slate": "💼 **Corporate Slate**: Gray-blue - Professional & Modern",
        "vibrant_education": "🎨 **Vibrant Education**: Multi-color - Colorful & Interactive",
        "modern_blue": "🔵 **Modern Blue**: Classic blue - Corporate Professional",
        "elegant_dark": "⚫ **Elegant Dark**: Dark theme - Premium & Sophisticated"
    }

    st.sidebar.info(template_descriptions.get(template_option, "Professional template"))

# Number of slides
num_slides = st.sidebar.slider(
    "Number of Slides",
    min_value=5,
    max_value=20,
    value=10,
    help="How many slides should the presentation have? (Educational topics work best with 8-12 slides)"
)

# Custom filename option
custom_filename = st.sidebar.text_input(
    "Output Filename (optional)",
    placeholder="my_presentation.pptx",
    help="Leave empty to auto-generate filename"
)

st.sidebar.markdown("---")
st.sidebar.markdown("### 💡 Tips")
st.sidebar.markdown("""
- Be specific with your topic
- Use clear, descriptive prompts
- More slides = more detailed content
- Generated presentations include speaker notes
""")

# ========== MAIN CONTENT AREA ==========

# Create two columns for better layout
col1, col2 = st.columns([2, 1])

with col1:
    st.header("📝 Enter Your Topic")

    # Topic input - main text area
    topic = st.text_area(
        "Enter Your Teaching Topic",
        placeholder="Example: 'Photosynthesis in Plants' or 'The Water Cycle and Its Importance'",
        height=150,
        help="Enter any educational topic - the AI will create a comprehensive, detailed presentation perfect for classroom teaching!"
    )

    # Generate button
    generate_button = st.button("🚀 Generate Detailed Presentation", type="primary", use_container_width=True)

with col2:
    st.header("📋 Example Topics")
    st.markdown("""
    **Science:**
    - Photosynthesis in Plants
    - Human Digestive System
    - Solar System and Planets
    - States of Matter

    **Mathematics:**
    - Introduction to Fractions
    - Pythagorean Theorem
    - Linear Equations

    **Social Studies:**
    - Indian Freedom Struggle
    - World War II
    - The Water Cycle

    **Languages:**
    - Parts of Speech
    - Essay Writing Techniques
    """)

# ========== GENERATION LOGIC ==========
# This runs when the user clicks the generate button
if generate_button:
    # Validate that a topic was entered
    if not topic or topic.strip() == "":
        st.error("❌ Please enter a topic for your presentation!")
    else:
        # Show progress and status messages
        with st.spinner("🤖 AI is generating your presentation... This may take 30-60 seconds..."):
            try:
                # Create a status container for updates
                status_container = st.empty()

                # Step 1: Show initial status
                status_container.info("🔄 Step 1/3: Contacting OpenAI API...")

                # Prepare filename based on output format
                if custom_filename and custom_filename.strip():
                    filename = custom_filename.strip()
                    # Add appropriate extension based on format
                    if is_html:
                        if not filename.endswith('.html'):
                            filename += '.html'
                    else:
                        if not filename.endswith('.pptx'):
                            filename += '.pptx'
                else:
                    filename = None  # Let the function auto-generate

                # Step 2: Update status
                status_container.info("🔄 Step 2/3: Generating slide content with AI...")

                # Import generation modules to access slide data
                from ppt_generator import initialize_openai_client, generate_slide_content, create_presentation_from_template
                from html_generator import create_html_presentation

                # Initialize OpenAI and generate slide content (only once)
                client = initialize_openai_client()
                slides_data = generate_slide_content(client, topic, num_slides)

                # Store in session state for potential conversion
                st.session_state['slides_data'] = slides_data
                st.session_state['topic'] = topic
                st.session_state['num_slides'] = num_slides

                # Call the appropriate generation function based on format
                if is_html:
                    # Prepare filename
                    if not filename:
                        safe_topic = "".join(c if c.isalnum() else "_" for c in topic)
                        safe_topic = safe_topic[:50]
                        filename = f"{safe_topic}_presentation.html"

                    output_path = create_html_presentation(slides_data, filename, template_option)
                    file_type = "HTML Presentation"
                    mime_type = "text/html"
                    download_label = "⬇️ Download HTML Presentation"

                    # Also generate PowerPoint version with same styling
                    status_container.info("🔄 Also creating PowerPoint version with same styling...")
                    pptx_filename = f"{safe_topic}_{template_option}.pptx"
                    from html_to_ppt_converter import convert_html_style_to_pptx
                    pptx_output_path = convert_html_style_to_pptx(slides_data, pptx_filename, template_option)

                    # Store both paths
                    st.session_state['html_path'] = output_path
                    st.session_state['pptx_path'] = pptx_output_path
                else:
                    # Prepare filename
                    if not filename:
                        safe_topic = "".join(c if c.isalnum() else "_" for c in topic)
                        safe_topic = safe_topic[:50]
                        filename = f"{safe_topic}_presentation.pptx"

                    output_path = create_presentation_from_template(template_option, slides_data, filename)
                    file_type = "PowerPoint"
                    mime_type = "application/vnd.openxmlformats-officedocument.presentationml.presentation"
                    download_label = "⬇️ Download PowerPoint"

                # Step 3: Update status
                status_container.info(f"🔄 Step 3/3: Creating {file_type} file...")

                # Success! Clear the status and show success message
                status_container.empty()
                st.success(f"✅ {file_type} presentation generated successfully!")

                # Display download button
                st.markdown("### 📥 Download Your Presentation")

                # Read the generated file
                with open(output_path, 'rb') as file:
                    file_data = file.read()

                # Extract just the filename for display
                display_filename = os.path.basename(output_path)

                # Create download buttons
                if is_html:
                    col_btn1, col_btn2, col_btn3 = st.columns(3)
                else:
                    col_btn1, col_btn2 = st.columns(2)

                with col_btn1:
                    st.download_button(
                        label=download_label,
                        data=file_data,
                        file_name=display_filename,
                        mime=mime_type,
                        type="primary",
                        use_container_width=True
                    )

                # Add "Open in Browser" button for HTML
                if is_html:
                    with col_btn2:
                        # Read HTML content for inline display
                        with open(output_path, 'r', encoding='utf-8') as f:
                            html_content = f.read()

                        # Create a button that opens HTML in new tab
                        import base64
                        b64 = base64.b64encode(html_content.encode()).decode()
                        href = f'<a href="data:text/html;base64,{b64}" download="{display_filename}" target="_blank"><button style="width:100%; padding:0.5rem 1rem; background-color:#10b981; color:white; border:none; border-radius:0.375rem; font-weight:600; cursor:pointer; font-size:1rem;">🌐 Open in New Tab</button></a>'
                        st.markdown(href, unsafe_allow_html=True)

                    # Add PowerPoint download button
                    with col_btn3:
                        if 'pptx_path' in st.session_state:
                            with open(st.session_state['pptx_path'], 'rb') as pptx_file:
                                pptx_data = pptx_file.read()

                            st.download_button(
                                label="📥 Download as .pptx",
                                data=pptx_data,
                                file_name=os.path.basename(st.session_state['pptx_path']),
                                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                                type="secondary",
                                use_container_width=True
                            )

                    st.info("💡 **Three ways to use your presentation:**\n"
                           "- 📥 **Download HTML** - Interactive version for browsers\n"
                           "- 🌐 **Open in New Tab** - View it now in your browser\n"
                           "- 📥 **Download as .pptx** - PowerPoint version with same styling\n"
                           "- Use arrow keys ← → or click buttons to navigate\n"
                           "- Press 'F' for fullscreen, 'N' for notes")

                    # Add inline preview option for HTML
                    with st.expander("👁️ Preview Presentation (Inline)"):
                        st.markdown("**Note:** For best experience, use '🌐 Open in New Tab' button above.")
                        # Display HTML inline using components
                        import streamlit.components.v1 as components
                        components.html(html_content, height=600, scrolling=True)

                # Show presentation details
                st.info(f"""
                **Presentation Details:**
                - Topic: {topic}
                - Format: {file_type}
                - Template/Style: {template_option}
                - Number of Slides: {num_slides}
                - File: {display_filename}
                """)

                # Show tips for using the presentation
                with st.expander("💡 Next Steps"):
                    if is_html:
                        st.markdown("""
                        1. **Download** the HTML file using the button above
                        2. **Open** it in any web browser (works offline!)
                        3. **Navigate** using arrow keys or navigation buttons
                        4. **Present** in fullscreen mode (press F key)
                        5. **View notes** by pressing N key during presentation
                        6. **Customize** the HTML/CSS if needed (it's just a single file!)

                        **Keyboard Shortcuts:**
                        - ← → : Previous/Next slide
                        - F : Toggle fullscreen
                        - N : Toggle speaker notes
                        - Space : Next slide
                        """)
                    else:
                        st.markdown("""
                        1. **Download** the presentation using the button above
                        2. **Open** it in Microsoft PowerPoint, Google Slides, or LibreOffice
                        3. **Review** the content and customize as needed
                        4. **Replace** image placeholders with actual images
                        5. **Edit** colors, fonts, and layouts to match your style
                        6. **Check** speaker notes for additional talking points

                        **Note:** The gray rectangles marked "[Image Placeholder]" are where you can add relevant images.
                        """)

            except Exception as e:
                # If any error occurs, show it to the user
                st.error(f"❌ Error generating presentation: {str(e)}")
                st.markdown("**Troubleshooting:**")
                st.markdown("""
                - Check your OpenAI API key in the .env file
                - Ensure you have sufficient API credits
                - Try a simpler topic or fewer slides
                - Check your internet connection
                """)

                # Show detailed error in an expander (for debugging)
                with st.expander("🔍 Technical Error Details"):
                    import traceback
                    st.code(traceback.format_exc())

# ========== FOOTER ==========
st.markdown("---")
st.markdown("""
<div style='text-align: center; color: gray;'>
    <p>🤖 Powered by OpenAI GPT | Built with Streamlit & python-pptx</p>
    <p>Made for creating professional presentations quickly and easily</p>
</div>
""", unsafe_allow_html=True)
