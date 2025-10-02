"""
HTML Presentation Generator with Tailwind CSS
Creates beautiful, interactive presentations using HTML and Tailwind CSS
"""

import os
import json
import base64
from ppt_generator import initialize_openai_client, generate_slide_content

# Import Pexels image fetcher
try:
    from images.pexels_fetcher import get_image_for_slide
    PEXELS_AVAILABLE = True
except ImportError:
    PEXELS_AVAILABLE = False
    print("Warning: Pexels image fetcher not available")


def create_html_presentation(slides_data, output_filename, template_style="vibrant"):
    """
    Creates an HTML presentation with Tailwind CSS from slide data

    Args:
        slides_data: List of slide dictionaries from generate_slide_content()
        output_filename: Name for output HTML file
        template_style: "vibrant", "modern", or "dark"

    Returns:
        Path to the created HTML file
    """

    # Color schemes for different templates
    templates = {
        "vibrant": {
            "primary": "bg-gradient-to-r from-purple-600 to-blue-600",
            "secondary": "bg-gradient-to-r from-green-400 to-blue-500",
            "text": "text-gray-800",
            "bg": "bg-white",
            "accent": "text-orange-500",
            "content_bg": "bg-gradient-to-br from-gray-50 to-gray-100",
            "card_bg": "bg-white/80",
            "border": "border-gray-200/50"
        },
        "modern": {
            "primary": "bg-gradient-to-r from-blue-600 to-indigo-700",
            "secondary": "bg-gradient-to-r from-blue-400 to-blue-600",
            "text": "text-gray-900",
            "bg": "bg-gray-50",
            "accent": "text-blue-600",
            "content_bg": "bg-gradient-to-br from-gray-50 to-gray-100",
            "card_bg": "bg-white/80",
            "border": "border-gray-200/50"
        },
        "dark": {
            "primary": "bg-gradient-to-r from-gray-900 to-gray-800",
            "secondary": "bg-gradient-to-r from-purple-900 to-indigo-900",
            "text": "text-gray-100",
            "bg": "bg-gray-900",
            "accent": "text-yellow-400",
            "content_bg": "bg-gradient-to-br from-gray-800 to-gray-900",
            "card_bg": "bg-gray-800/80",
            "border": "border-gray-600/50"
        }
    }

    theme = templates.get(template_style, templates["vibrant"])

    # Generate HTML content
    html_content = f"""<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>{slides_data[0]['title'] if slides_data else 'Presentation'}</title>
    <script src="https://cdn.tailwindcss.com"></script>
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&family=Poppins:wght@600;700;800&display=swap" rel="stylesheet">
    <style>
        * {{
            font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif;
        }}
        h1, h2, h3 {{
            font-family: 'Poppins', sans-serif;
        }}
        .slide {{
            height: 100vh;
            display: none;
            opacity: 0;
            transform: scale(0.95);
            transition: all 0.4s cubic-bezier(0.4, 0, 0.2, 1);
            overflow: hidden;
        }}
        .slide.active {{
            display: flex;
            opacity: 1;
            transform: scale(1);
        }}
        @keyframes slideInRight {{
            from {{
                opacity: 0;
                transform: translateX(30px);
            }}
            to {{
                opacity: 1;
                transform: translateX(0);
            }}
        }}
        @keyframes slideInLeft {{
            from {{
                opacity: 0;
                transform: translateX(-30px);
            }}
            to {{
                opacity: 1;
                transform: translateX(0);
            }}
        }}
        @keyframes fadeInUp {{
            from {{
                opacity: 0;
                transform: translateY(30px);
            }}
            to {{
                opacity: 1;
                transform: translateY(0);
            }}
        }}
        @keyframes scaleIn {{
            from {{
                opacity: 0;
                transform: scale(0.8);
            }}
            to {{
                opacity: 1;
                transform: scale(1);
            }}
        }}
        @keyframes shimmer {{
            0% {{ background-position: -1000px 0; }}
            100% {{ background-position: 1000px 0; }}
        }}
        .fade-in {{
            animation: fadeInUp 0.6s ease-out;
        }}
        .slide-in-right {{
            animation: slideInRight 0.7s ease-out;
        }}
        .slide-in-left {{
            animation: slideInLeft 0.7s ease-out;
        }}
        .scale-in {{
            animation: scaleIn 0.5s ease-out;
        }}
        .bullet-point {{
            position: relative;
            padding-left: 2rem;
            transition: all 0.3s ease;
        }}
        .bullet-point:hover {{
            transform: translateX(8px);
            background: linear-gradient(90deg, transparent, rgba(249, 115, 22, 0.1));
            border-radius: 8px;
            padding: 0.5rem 0.5rem 0.5rem 2rem;
        }}
        .bullet-point::before {{
            content: "";
            position: absolute;
            left: 0;
            top: 0.6rem;
            width: 12px;
            height: 12px;
            background: linear-gradient(135deg, #f97316, #fb923c);
            border-radius: 50%;
            box-shadow: 0 2px 8px rgba(249, 115, 22, 0.4);
        }}
        .gradient-text {{
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
            background-clip: text;
        }}
        .card-3d {{
            transform-style: preserve-3d;
            transition: transform 0.3s ease;
        }}
        .card-3d:hover {{
            transform: translateY(-5px) rotateX(5deg);
        }}
        .shimmer {{
            background: linear-gradient(90deg, transparent, rgba(255,255,255,0.3), transparent);
            background-size: 200% 100%;
            animation: shimmer 3s infinite;
        }}
        .progress-bar {{
            position: fixed;
            top: 0;
            left: 0;
            height: 4px;
            background: linear-gradient(90deg, #667eea, #764ba2, #f093fb, #4facfe);
            transition: width 0.3s ease;
            z-index: 9999;
        }}
        .blob {{
            border-radius: 30% 70% 70% 30% / 30% 30% 70% 70%;
            animation: blob 8s infinite;
        }}
        @keyframes blob {{
            0%, 100% {{ border-radius: 30% 70% 70% 30% / 30% 30% 70% 70%; }}
            25% {{ border-radius: 58% 42% 75% 25% / 76% 46% 54% 24%; }}
            50% {{ border-radius: 50% 50% 33% 67% / 55% 27% 73% 45%; }}
            75% {{ border-radius: 33% 67% 58% 42% / 63% 68% 32% 37%; }}
        }}
        .glass {{
            background: rgba(255, 255, 255, 0.1);
            backdrop-filter: blur(10px);
            border: 1px solid rgba(255, 255, 255, 0.2);
        }}
    </style>
</head>
<body class="{theme['bg']} {theme['text']}">{"""
    <!-- Progress Bar -->
    <div class="progress-bar" id="progressBar"></div>"""}
    <div id="presentation" class="relative">
"""

    # Generate slides
    for idx, slide in enumerate(slides_data):
        is_title = idx == 0

        if is_title:
            # Title Slide - Enhanced with modern styling
            html_content += f"""
        <div class="slide {'active' if idx == 0 else ''}" data-slide="{idx}">
            <div class="w-full h-full {theme['primary']} flex items-center justify-center p-8 relative overflow-hidden">
                <!-- Animated background blobs -->
                <div class="absolute top-10 left-10 w-72 h-72 bg-white/10 rounded-full blob blur-3xl"></div>
                <div class="absolute bottom-10 right-10 w-96 h-96 bg-white/10 rounded-full blob blur-3xl" style="animation-delay: -4s;"></div>

                <!-- Content -->
                <div class="text-center space-y-6 fade-in relative z-10">
                    <div class="mb-4">
                        <div class="inline-block px-5 py-2 bg-white/20 backdrop-blur-md rounded-full text-white/90 text-sm font-semibold mb-6 shimmer">
                            ✨ AI-Powered Presentation
                        </div>
                    </div>

                    <h1 class="text-6xl font-extrabold text-white mb-6 leading-tight scale-in px-4">
                        {slide['title']}
                    </h1>

                    <div class="h-1 w-32 mx-auto bg-gradient-to-r from-transparent via-white to-transparent rounded-full"></div>

                    <p class="text-2xl text-white/90 font-light max-w-3xl mx-auto slide-in-right px-4">
                        {slide['content'][0] if slide.get('content') else ''}
                    </p>

                    <div class="mt-12 slide-in-left">
                        <div class="inline-flex items-center space-x-3 px-6 py-3 glass rounded-2xl shadow-2xl">
                            <svg class="w-8 h-8 text-white" fill="currentColor" viewBox="0 0 20 20">
                                <path d="M10.394 2.08a1 1 0 00-.788 0l-7 3a1 1 0 000 1.84L5.25 8.051a.999.999 0 01.356-.257l4-1.714a1 1 0 11.788 1.838L7.667 9.088l1.94.831a1 1 0 00.787 0l7-3a1 1 0 000-1.838l-7-3zM3.31 9.397L5 10.12v4.102a8.969 8.969 0 00-1.05-.174 1 1 0 01-.89-.89 11.115 11.115 0 01.25-3.762zM9.3 16.573A9.026 9.026 0 007 14.935v-3.957l1.818.78a3 3 0 002.364 0l5.508-2.361a11.026 11.026 0 01.25 3.762 1 1 0 01-.89.89 8.968 8.968 0 00-5.35 2.524 1 1 0 01-1.4 0zM6 18a1 1 0 001-1v-2.065a8.935 8.935 0 00-2-.712V17a1 1 0 001 1z"/>
                            </svg>
                            <span class="text-lg font-semibold text-white">Interactive Learning Experience</span>
                        </div>
                    </div>
                </div>
            </div>
        </div>
"""
        else:
            # Content Slide
            has_image = slide.get('has_image', False)
            content_points = slide.get('content', [])

            # Process bullet points to extract bold text
            formatted_points = []
            for point in content_points:
                # Convert **text** to <strong>text</strong>
                import re
                formatted = re.sub(r'\*\*(.*?)\*\*', r'<strong class="font-bold ' + theme['accent'] + r'">\1</strong>', point)
                formatted_points.append(formatted)

            html_content += f"""
        <div class="slide" data-slide="{idx}">
            <div class="w-full h-full p-8 {theme['content_bg']} overflow-auto">
                <!-- Header with gradient and shadow -->
                <div class="{theme['primary']} text-white p-6 rounded-2xl shadow-2xl mb-6 card-3d relative overflow-hidden">
                    <div class="absolute top-0 right-0 w-64 h-64 bg-white/10 rounded-full blur-3xl"></div>
                    <h2 class="text-4xl font-bold relative z-10 slide-in-left">{slide['title']}</h2>
                    <div class="mt-3 h-1.5 w-24 bg-white/40 rounded-full relative z-10"></div>
                </div>

                <!-- Content with enhanced cards -->
                <div class="{theme['card_bg']} backdrop-blur-sm border {theme['border']} rounded-2xl p-6 shadow-xl">
                    <div class="{'grid grid-cols-2 gap-8' if has_image else ''}">
                        <div class="space-y-4">
"""

            for i, point in enumerate(formatted_points):
                delay = i * 0.1
                html_content += f"""
                            <div class="bullet-point text-lg leading-relaxed {theme['text']} fade-in p-3 rounded-xl shadow-sm hover:shadow-md transition-all duration-300" style="animation-delay: {delay}s;">
                                {point}
                            </div>
"""

            html_content += """
                        </div>
"""

            if has_image:
                # Try to get real image from Pexels
                image_html = ""
                image_added = False

                if PEXELS_AVAILABLE:
                    try:
                        # Get image for this slide
                        image_path = get_image_for_slide(slide['title'], idx)

                        if image_path and os.path.exists(image_path):
                            # Read image and convert to base64
                            with open(image_path, 'rb') as img_file:
                                img_data = base64.b64encode(img_file.read()).decode()

                            # Create image HTML with base64 embedded image
                            image_html = f"""
                        <div class="flex items-center justify-center slide-in-right">
                            <img src="data:image/jpeg;base64,{img_data}"
                                 alt="{slide['title']}"
                                 class="w-full h-80 object-cover rounded-2xl shadow-xl card-3d">
                        </div>
"""
                            image_added = True
                            print(f"  ✓ Embedded Pexels image in HTML for slide {idx+1}")
                    except Exception as e:
                        print(f"  ⚠ Could not embed Pexels image in HTML: {e}")

                # Fallback to placeholder if no image
                if not image_added:
                    image_html = f"""
                        <div class="flex items-center justify-center slide-in-right">
                            <div class="w-full h-80 bg-gradient-to-br from-indigo-100 via-purple-100 to-pink-100 rounded-2xl flex items-center justify-center border-4 border-dashed border-indigo-300 shadow-xl card-3d relative overflow-hidden">
                                <div class="absolute inset-0 bg-gradient-to-r from-transparent via-white/30 to-transparent shimmer"></div>
                                <div class="text-center relative z-10">
                                    <svg class="w-24 h-24 mx-auto text-indigo-400 mb-3" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                                        <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M4 16l4.586-4.586a2 2 0 012.828 0L16 16m-2-2l1.586-1.586a2 2 0 012.828 0L20 14m-6-6h.01M6 20h12a2 2 0 002-2V6a2 2 0 00-2-2H6a2 2 0 00-2 2v12a2 2 0 002 2z"></path>
                                    </svg>
                                    <p class="text-indigo-600 font-bold text-base">📸 Visual Content</p>
                                    <p class="text-indigo-400 text-sm mt-1">Add image or diagram</p>
                                </div>
                            </div>
                        </div>
"""

                html_content += image_html

            html_content += f"""
                    </div>
                </div>

                <!-- Speaker Notes with modern design -->
                <div class="mt-4 p-4 bg-gradient-to-r from-blue-500 to-indigo-600 text-white border-l-4 border-yellow-400 rounded-xl shadow-xl hidden" id="notes-{idx}">
                    <div class="flex items-start space-x-3">
                        <svg class="w-6 h-6 flex-shrink-0 mt-1" fill="currentColor" viewBox="0 0 20 20">
                            <path d="M9 2a1 1 0 000 2h2a1 1 0 100-2H9z"></path>
                            <path fill-rule="evenodd" d="M4 5a2 2 0 012-2 3 3 0 003 3h2a3 3 0 003-3 2 2 0 012 2v11a2 2 0 01-2 2H6a2 2 0 01-2-2V5zm3 4a1 1 0 000 2h.01a1 1 0 100-2H7zm3 0a1 1 0 000 2h3a1 1 0 100-2h-3zm-3 4a1 1 0 100 2h.01a1 1 0 100-2H7zm3 0a1 1 0 100 2h3a1 1 0 100-2h-3z" clip-rule="evenodd"></path>
                        </svg>
                        <div class="flex-1">
                            <h3 class="font-bold text-base mb-2 flex items-center">
                                <span class="mr-2">📝</span> Teacher's Notes
                            </h3>
                            <p class="text-white/95 text-sm leading-relaxed">{slide.get('notes', '')}</p>
                        </div>
                    </div>
                </div>
            </div>
        </div>
"""

    # Add navigation and controls
    html_content += f"""
    </div>

    <!-- Enhanced Navigation Controls -->
    <div class="fixed bottom-8 left-1/2 transform -translate-x-1/2 flex items-center space-x-3 bg-gradient-to-r from-white/95 to-gray-50/95 backdrop-blur-md rounded-full px-8 py-4 shadow-2xl border border-gray-200/50 scale-in">
        <!-- Previous Button -->
        <button onclick="prevSlide()" class="group p-3 hover:bg-gradient-to-r hover:from-purple-500 hover:to-indigo-600 rounded-full transition-all duration-300 hover:scale-110 hover:shadow-lg">
            <svg class="w-6 h-6 text-gray-700 group-hover:text-white transition" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2.5" d="M15 19l-7-7 7-7"></path>
            </svg>
        </button>

        <!-- Slide Counter with gradient background -->
        <div class="flex items-center space-x-3 px-5 py-2 bg-gradient-to-r from-purple-500 to-indigo-600 rounded-full">
            <span id="currentSlide" class="font-bold text-white text-lg">1</span>
            <span class="text-white/60 font-bold">/</span>
            <span id="totalSlides" class="text-white/90 font-semibold">{len(slides_data)}</span>
        </div>

        <!-- Next Button -->
        <button onclick="nextSlide()" class="group p-3 hover:bg-gradient-to-r hover:from-purple-500 hover:to-indigo-600 rounded-full transition-all duration-300 hover:scale-110 hover:shadow-lg">
            <svg class="w-6 h-6 text-gray-700 group-hover:text-white transition" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2.5" d="M9 5l7 7-7 7"></path>
            </svg>
        </button>

        <!-- Divider -->
        <div class="h-8 w-px bg-gray-300 mx-2"></div>

        <!-- Notes Button -->
        <button onclick="toggleNotes()" class="group p-3 hover:bg-gradient-to-r hover:from-blue-500 hover:to-cyan-600 rounded-full transition-all duration-300 hover:scale-110 hover:shadow-lg" title="Toggle Notes (N)">
            <svg class="w-6 h-6 text-gray-700 group-hover:text-white transition" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 12h6m-6 4h6m2 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z"></path>
            </svg>
        </button>

        <!-- Fullscreen Button -->
        <button onclick="toggleFullscreen()" class="group p-3 hover:bg-gradient-to-r hover:from-green-500 hover:to-emerald-600 rounded-full transition-all duration-300 hover:scale-110 hover:shadow-lg" title="Fullscreen (F)">
            <svg class="w-6 h-6 text-gray-700 group-hover:text-white transition" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M4 8V4m0 0h4M4 4l5 5m11-1V4m0 0h-4m4 0l-5 5M4 16v4m0 0h4m-4 0l5-5m11 5l-5-5m5 5v-4m0 4h-4"></path>
            </svg>
        </button>
    </div>

    <!-- Keyboard Shortcuts Indicator -->
    <div class="fixed top-6 right-6 bg-white/90 backdrop-blur-md rounded-2xl shadow-xl p-4 border border-gray-200/50 hidden" id="shortcutsPanel">
        <h3 class="font-bold text-gray-800 mb-3 flex items-center">
            <svg class="w-5 h-5 mr-2" fill="currentColor" viewBox="0 0 20 20">
                <path fill-rule="evenodd" d="M18 10a8 8 0 11-16 0 8 8 0 0116 0zm-8-3a1 1 0 00-.867.5 1 1 0 11-1.731-1A3 3 0 0113 8a3.001 3.001 0 01-2 2.83V11a1 1 0 11-2 0v-1a1 1 0 011-1 1 1 0 100-2zm0 8a1 1 0 100-2 1 1 0 000 2z" clip-rule="evenodd"></path>
            </svg>
            Keyboard Shortcuts
        </h3>
        <div class="space-y-2 text-sm text-gray-700">
            <div class="flex justify-between"><kbd class="px-2 py-1 bg-gray-100 rounded">←→</kbd> <span>Navigate</span></div>
            <div class="flex justify-between"><kbd class="px-2 py-1 bg-gray-100 rounded">Space</kbd> <span>Next slide</span></div>
            <div class="flex justify-between"><kbd class="px-2 py-1 bg-gray-100 rounded">N</kbd> <span>Toggle notes</span></div>
            <div class="flex justify-between"><kbd class="px-2 py-1 bg-gray-100 rounded">F</kbd> <span>Fullscreen</span></div>
        </div>
    </div>

    <!-- Help Button -->
    <button onclick="toggleShortcuts()" class="fixed top-6 right-6 p-3 bg-gradient-to-r from-purple-500 to-indigo-600 text-white rounded-full shadow-lg hover:shadow-xl transition-all duration-300 hover:scale-110">
        <svg class="w-6 h-6" fill="currentColor" viewBox="0 0 20 20">
            <path fill-rule="evenodd" d="M18 10a8 8 0 11-16 0 8 8 0 0116 0zm-8-3a1 1 0 00-.867.5 1 1 0 11-1.731-1A3 3 0 0113 8a3.001 3.001 0 01-2 2.83V11a1 1 0 11-2 0v-1a1 1 0 011-1 1 1 0 100-2zm0 8a1 1 0 100-2 1 1 0 000 2z" clip-rule="evenodd"></path>
        </svg>
    </button>

    <script>
        let currentSlide = 0;
        const slides = document.querySelectorAll('.slide');
        const totalSlides = slides.length;

        document.getElementById('totalSlides').textContent = totalSlides;

        // Initialize progress bar
        showSlide(0);

        function showSlide(index) {{
            slides.forEach(slide => slide.classList.remove('active'));
            slides[index].classList.add('active');
            currentSlide = index;
            document.getElementById('currentSlide').textContent = index + 1;

            // Update progress bar
            const progress = ((index + 1) / totalSlides) * 100;
            document.getElementById('progressBar').style.width = progress + '%';
        }}

        function nextSlide() {{
            if (currentSlide < totalSlides - 1) {{
                showSlide(currentSlide + 1);
            }}
        }}

        function prevSlide() {{
            if (currentSlide > 0) {{
                showSlide(currentSlide - 1);
            }}
        }}

        function toggleNotes() {{
            const notes = document.getElementById('notes-' + currentSlide);
            if (notes) {{
                notes.classList.toggle('hidden');
            }}
        }}

        function toggleFullscreen() {{
            if (!document.fullscreenElement) {{
                document.documentElement.requestFullscreen();
            }} else {{
                document.exitFullscreen();
            }}
        }}

        function toggleShortcuts() {{
            const panel = document.getElementById('shortcutsPanel');
            panel.classList.toggle('hidden');
        }}

        // Keyboard navigation
        document.addEventListener('keydown', (e) => {{
            if (e.key === 'ArrowRight' || e.key === ' ') {{
                nextSlide();
            }} else if (e.key === 'ArrowLeft') {{
                prevSlide();
            }} else if (e.key === 'n' || e.key === 'N') {{
                toggleNotes();
            }} else if (e.key === 'f' || e.key === 'F') {{
                toggleFullscreen();
            }}
        }});
    </script>
</body>
</html>
"""

    # Save HTML file
    os.makedirs('output', exist_ok=True)
    output_path = f"output/{output_filename}"

    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(html_content)

    print(f"HTML presentation saved: {output_path}")
    return output_path


def generate_html_ppt(topic, template_style="vibrant", num_slides=10, output_filename=None):
    """
    Main function to generate HTML presentation with Tailwind CSS

    Args:
        topic: The topic/prompt for the presentation
        template_style: "vibrant", "modern", or "dark"
        num_slides: Number of slides to generate
        output_filename: Name for output HTML file (auto-generated if None)

    Returns:
        Path to the generated HTML file
    """

    # Auto-generate filename if not provided
    if output_filename is None:
        safe_topic = "".join(c if c.isalnum() else "_" for c in topic)
        safe_topic = safe_topic[:50]
        output_filename = f"{safe_topic}_presentation.html"

    # Step 1: Initialize OpenAI client
    print("Step 1: Initializing OpenAI client...")
    client = initialize_openai_client()

    # Step 2: Generate slide content using AI
    print(f"\nStep 2: Generating content for {num_slides} slides on topic: '{topic}'")
    slides_data = generate_slide_content(client, topic, num_slides)

    # Step 3: Create the HTML presentation
    print(f"\nStep 3: Creating HTML presentation with style '{template_style}'")
    output_path = create_html_presentation(slides_data, output_filename, template_style)

    print("\n[SUCCESS] HTML presentation generated successfully!")
    return output_path


# Test function
if __name__ == "__main__":
    print("HTML Presentation Generator - Test Mode")
    print("=" * 50)

    test_topic = "Photosynthesis in Plants"
    test_template = "vibrant"
    test_slides = 8

    try:
        result_path = generate_html_ppt(
            topic=test_topic,
            template_style=test_template,
            num_slides=test_slides
        )
        print(f"\nTest successful! Open your presentation at: {result_path}")
    except Exception as e:
        print(f"\nError during generation: {e}")
        import traceback
        traceback.print_exc()
