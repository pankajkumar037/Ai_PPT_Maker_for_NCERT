"""
PPTMaker-X Streamlit App (Gamma.AI Style)
Build each slide in PPT, show preview, get feedback, move to next
"""

import streamlit as st
import os
from datetime import datetime
from pptmaker_x import ConversationalPPTAgent, TemplatePPTGenerator

# Page configuration
st.set_page_config(
    page_title="PPTMaker-X",
    page_icon="🤖",
    layout="wide"
)

# Custom CSS
st.markdown("""
<style>
    .main-header {
        font-size: 3rem;
        font-weight: bold;
        text-align: center;
        margin-bottom: 0;
    }
    .sub-header {
        text-align: center;
        color: #666;
        margin-bottom: 2rem;
    }
    .slide-preview {
        border: 2px solid #0066CC;
        border-radius: 10px;
        padding: 1rem;
        margin: 1rem 0;
    }
</style>
""", unsafe_allow_html=True)

# Header
st.markdown('<div class="main-header">🤖 PPTMaker-X</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-header">AI Presentation Generator (Gamma.AI Style)</div>', unsafe_allow_html=True)

# Initialize session state
if 'agent' not in st.session_state:
    st.session_state.agent = None
if 'outline' not in st.session_state:
    st.session_state.outline = None
if 'current_slide_index' not in st.session_state:
    st.session_state.current_slide_index = 0
if 'ppt_path' not in st.session_state:
    st.session_state.ppt_path = None
if 'step' not in st.session_state:
    st.session_state.step = 'input'  # input -> outline -> slides -> done
if 'total_slides' not in st.session_state:
    st.session_state.total_slides = 0
if 'filename' not in st.session_state:
    st.session_state.filename = None

# STEP 1: Input
if st.session_state.step == 'input':
    with st.expander("ℹ️ How It Works", expanded=False):
        st.markdown("""
        **Gamma.AI Style Workflow with AutoGen Agent:**

        1. **📋 Generate Outline** - AutoGen agent creates presentation structure
        2. **🎨 Select Theme** - AutoGen agent picks the best template
        3. **For Each Slide:**
           - ✍️ AutoGen agent generates content
           - 📊 Builds slide directly in PowerPoint file
           - 📄 Shows you the ACTUAL PPTX slide content
           - 💬 You give feedback or approve
           - 🔄 AutoGen agent modifies slide IN-PLACE (no regeneration)
        4. **📥 Download** - Get your final presentation

        **Key Features:**
        - ✅ Direct PPTX rendering (what you see = what you download)
        - ✅ In-place slide modification (not regeneration)
        - ✅ AutoGen conversational agent
        - ✅ You approve EACH SLIDE before moving to next
        """)

    st.markdown("### 📋 Enter Presentation Details")

    topic = st.text_area(
        "Presentation Topic",
        placeholder="Example: Create a presentation about Renewable Energy",
        height=100
    )

    col1, col2 = st.columns(2)
    with col1:
        num_slides = st.number_input(
            "Number of Slides",
            min_value=3,
            max_value=15,
            value=6
        )

    with col2:
        st.markdown("###")
        st.info(f"📊 Will generate **{num_slides} slides**")

    if st.button("🚀 Start Generation", use_container_width=True, type="primary"):
        if not topic.strip():
            st.error("⚠️ Please enter a topic!")
        elif not os.getenv("OPENAI_API_KEY"):
            st.error("⚠️ OpenAI API key not found in .env file!")
        else:
            # Initialize agent
            st.session_state.agent = ConversationalPPTAgent()
            st.session_state.topic = topic
            st.session_state.total_slides = num_slides
            st.session_state.current_slide_index = 0

            # Create filename
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            safe_topic = "".join(c if c.isalnum() else "_" for c in topic[:30])
            st.session_state.filename = f"{safe_topic}_{timestamp}"
            st.session_state.ppt_path = f"output/{st.session_state.filename}.pptx"

            st.session_state.step = 'outline'
            st.rerun()

# STEP 2: Generate Outline
elif st.session_state.step == 'outline':
    st.markdown("## 📋 Step 1: Generating Outline")

    with st.spinner("🤖 AI is creating presentation outline..."):
        try:
            outline = st.session_state.agent.generate_outline(
                st.session_state.topic,
                st.session_state.total_slides
            )
            st.session_state.outline = outline

            # Select theme
            theme = st.session_state.agent.select_theme()

            st.success(f"✅ Outline created! Using theme: **{theme.replace('_', ' ').title()}**")

            # Show outline
            st.markdown("### Presentation Outline:")
            for slide in outline.get('slides', []):
                st.markdown(f"**Slide {slide['number']}:** {slide['topic']} ({slide['type']})")

            st.markdown("---")
            if st.button("✅ Start Building Slides", use_container_width=True, type="primary"):
                st.session_state.step = 'slides'
                st.rerun()

        except Exception as e:
            st.error(f"❌ Error: {e}")
            if st.button("🔄 Restart"):
                st.session_state.step = 'input'
                st.rerun()

# STEP 3: Build slides one by one (Gamma.AI style)
elif st.session_state.step == 'slides':
    current_idx = st.session_state.current_slide_index
    total = st.session_state.total_slides

    # Progress bar
    progress = (current_idx) / total
    st.progress(progress)
    st.markdown(f"### Slide {current_idx + 1} of {total}")

    # Check if we're done with all slides
    if current_idx >= total:
        st.session_state.step = 'done'
        st.rerun()

    # Generate content for current slide (if not already generated)
    if current_idx >= len(st.session_state.agent.slides_content):
        with st.spinner(f"✍️ AI is generating content for slide {current_idx + 1}..."):
            try:
                slide_info = st.session_state.outline['slides'][current_idx]
                content = st.session_state.agent._generate_slide_content(slide_info)
                st.session_state.agent.slides_content.append(content)
            except Exception as e:
                st.error(f"❌ Error generating content: {e}")
                st.stop()

    # Build the slide in PowerPoint
    with st.spinner(f"📊 Building slide {current_idx + 1} in PowerPoint..."):
        try:
            ppt_path = st.session_state.agent.build_single_slide(
                current_idx,
                st.session_state.ppt_path
            )
        except Exception as e:
            st.error(f"❌ Error building slide: {e}")
            st.stop()

    # Show slide content
    slide_content = st.session_state.agent.slides_content[current_idx]

    col1, col2 = st.columns([1, 1])

    with col1:
        st.markdown("#### 📊 Actual PPTX Slide Content")
        st.info("This is the actual content in your PowerPoint slide")

        # Show actual PPTX slide content (no PNG conversion)
        with st.container():
            st.markdown(f"**Slide Type:** `{slide_content.get('slide_type', 'content')}`")

            if 'title' in slide_content:
                st.markdown(f"### {slide_content['title']}")

            if 'subtitle' in slide_content:
                st.markdown(f"*{slide_content['subtitle']}*")

            if 'bullets' in slide_content:
                st.markdown("**Content:**")
                for bullet in slide_content['bullets']:
                    st.markdown(f"• {bullet}")

            if 'text' in slide_content:
                st.markdown(f"**Text:** {slide_content['text']}")

            if 'stat' in slide_content:
                st.markdown(f"## {slide_content['stat']}")
                if 'description' in slide_content:
                    st.markdown(f"*{slide_content['description']}*")

        # Offer to view actual PPTX file
        st.markdown("---")
        with open(ppt_path, 'rb') as f:
            st.download_button(
                label="📥 View in PowerPoint",
                data=f.read(),
                file_name=f"preview_slide_{current_idx+1}.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                help="Download to view actual slide in PowerPoint",
                key=f"preview_{current_idx}"
            )

    with col2:
        st.markdown("#### 💬 Provide Feedback")

        feedback = st.text_area(
            "Suggest changes for this slide:",
            placeholder="e.g., 'Change title to...' or 'Add a bullet point about...'",
            help="AutoGen agent will modify the PPTX slide directly based on your feedback",
            key=f"feedback_{current_idx}"
        )

        st.markdown("---")

        col_a, col_b = st.columns(2)

        with col_a:
            if st.button("✅ Approve, Next Slide", use_container_width=True, type="primary"):
                st.session_state.current_slide_index += 1
                st.rerun()

        with col_b:
            if st.button("🔄 Apply Changes", use_container_width=True, disabled=not feedback.strip()):
                if feedback.strip():
                    with st.spinner("AutoGen agent modifying slide in-place..."):
                        try:
                            # AutoGen agent modifies slide directly in PPTX
                            updated_content = st.session_state.agent.modify_slide(
                                current_idx,
                                feedback,
                                st.session_state.ppt_path
                            )

                            st.success("✅ Slide modified in PPTX file!")
                            st.rerun()
                        except Exception as e:
                            st.error(f"Error: {e}")

# STEP 4: Done
elif st.session_state.step == 'done':
    st.success("✅ All slides completed!")

    st.markdown("### 📥 Download Your Presentation")

    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("Total Slides", st.session_state.total_slides)
    with col2:
        st.metric("Theme", st.session_state.agent.theme.replace('_', ' ').title())
    with col3:
        st.metric("Status", "Ready")

    # Download button
    ppt_path = st.session_state.ppt_path
    ppt_filename = f"{st.session_state.filename}.pptx"

    if os.path.exists(ppt_path):
        with open(ppt_path, 'rb') as f:
            st.download_button(
                label="📊 Download PowerPoint",
                data=f.read(),
                file_name=ppt_filename,
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                use_container_width=True,
                type="primary"
            )

    st.markdown("---")
    if st.button("🔄 Create Another Presentation", use_container_width=True):
        # Reset everything
        st.session_state.agent = None
        st.session_state.outline = None
        st.session_state.current_slide_index = 0
        st.session_state.ppt_path = None
        st.session_state.step = 'input'
        st.session_state.total_slides = 0
        st.session_state.filename = None
        st.rerun()

# Sidebar
with st.sidebar:
    st.markdown("### ℹ️ Current Status")

    if st.session_state.step == 'input':
        st.info("Waiting for input...")
    elif st.session_state.step == 'outline':
        st.info("Generating outline...")
    elif st.session_state.step == 'slides':
        st.info(f"Building slide {st.session_state.current_slide_index + 1}/{st.session_state.total_slides}")
    elif st.session_state.step == 'done':
        st.success("Complete!")

    st.markdown("---")
    st.markdown("### 🎨 Available Themes")
    st.markdown("""
    - Modern Blue
    - Elegant Purple
    - Corporate Green
    - Vibrant Orange
    - Dark Professional
    - Minimal Gray
    - Ocean Teal
    - Sunset Red
    - Royal Indigo
    - Forest Brown
    """)
