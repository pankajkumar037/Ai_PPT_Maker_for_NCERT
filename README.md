# 📊 AI-Powered PPT Maker

An intelligent PowerPoint presentation generator that uses OpenAI's GPT to create professional, well-structured presentations from simple text prompts.

## 🌟 Features

- **AI-Powered Content Generation**: Uses OpenAI GPT to generate slide titles, bullet points, and speaker notes
- **Professional Templates**: Choose from 2 beautifully designed templates:
  - Modern Blue (Corporate Professional)
  - Elegant Dark (Premium Modern)
- **Customizable**: Select number of slides (3-15)
- **Image Placeholders**: Automatically adds blank rectangles where images should go
- **Speaker Notes**: Each slide includes AI-generated speaker notes
- **Simple Interface**: User-friendly Streamlit web app
- **Beginner-Friendly**: Well-commented code, easy to understand and modify

## 📁 Project Structure

```
Ai_PPT_Maker_for_NCERT/
│
├── app.py                      # Streamlit web application
├── ppt_generator.py            # Main PPT generation logic with OpenAI integration
├── create_templates.py         # Script to generate template files
├── requirements.txt            # Python dependencies
├── .env.example                # Example environment variables file
├── README.md                   # This file
│
├── templates/                  # PPT template files
│   ├── modern_blue.pptx
│   └── elegant_dark.pptx
│
└── output/                     # Generated presentations (created automatically)
```

## 🚀 Setup Instructions

### 1. Prerequisites

- Python 3.8 or higher
- OpenAI API key ([Get one here](https://platform.openai.com/api-keys))

### 2. Installation

```bash
# Install required packages
pip install -r requirements.txt
```

### 3. Configure OpenAI API Key

Create a `.env` file in the project directory:

```bash
# Copy the example file
copy .env.example .env

# Edit .env and add your OpenAI API key:
OPENAI_API_KEY=your_actual_api_key_here
```

### 4. Generate Templates (First Time Only)

```bash
python create_templates.py
```

This creates two professional templates in the `templates/` folder.

## 🎯 How to Use

### Running the Web App

```bash
streamlit run app.py
```

This will open the web interface in your browser (usually at http://localhost:8501)

### Using the Web Interface

1. **Enter Your Topic**: Type what you want your presentation to be about
2. **Choose Settings** (sidebar):
   - Select a template (Modern Blue or Elegant Dark)
   - Choose number of slides (3-15)
   - Optionally set a custom filename
3. **Click "Generate Presentation"**: Wait 30-60 seconds for AI to work
4. **Download**: Click the download button to get your PowerPoint file

### Using the Code Directly (Advanced)

You can also use the PPT generator in your own Python scripts:

```python
from ppt_generator import generate_ppt

# Generate a presentation
output_path = generate_ppt(
    topic="Artificial Intelligence in Education",
    template_name="modern_blue",  # or "elegant_dark"
    num_slides=5,
    output_filename="my_presentation.pptx"
)

print(f"Presentation saved to: {output_path}")
```

## 🎨 Available Templates

### 1. Modern Blue
- **Style**: Professional corporate design
- **Colors**: Blue theme with white background
- **Best For**: Business presentations, formal reports, educational content

### 2. Elegant Dark
- **Style**: Premium modern design
- **Colors**: Dark theme with gold accents
- **Best For**: Tech presentations, creative pitches, modern topics

## 📝 Code Overview

### `ppt_generator.py` - Main Logic

**Key Functions:**

- `initialize_openai_client()`: Sets up OpenAI API connection
- `generate_slide_content()`: Calls OpenAI to generate structured content
- `create_presentation_from_template()`: Creates PPT slides from template and content
- `generate_ppt()`: Main function that orchestrates the entire process

**How It Works:**
1. Loads your OpenAI API key from `.env`
2. Sends your topic to OpenAI with a structured prompt
3. Receives JSON response with slide titles, bullet points, and notes
4. Creates slides using selected template
5. Adds content with proper formatting and colors
6. Inserts blank rectangles for image placeholders
7. Saves the completed presentation

### `app.py` - Streamlit Interface

- User-friendly web interface
- Template selection and settings
- Progress tracking during generation
- Download functionality
- Error handling and user feedback

### `create_templates.py` - Template Generator

- Creates professional PowerPoint templates
- Defines color schemes and layouts
- Generates `.pptx` files with styled backgrounds

## 🔧 Customization

### Adding More Templates

1. Edit `create_templates.py`
2. Add a new section to create your template
3. Define colors, shapes, and layouts
4. Run the script to generate the new template
5. Update `app.py` to include the new template in the dropdown

### Modifying AI Prompt

Edit the `prompt` variable in `ppt_generator.py` (line ~51) to change how AI generates content.

### Changing Slide Layouts

Modify the slide creation logic in `create_presentation_from_template()` function to adjust:
- Text positioning
- Font sizes and colors
- Image placeholder size and position
- Background elements

## 💡 Tips for Best Results

1. **Be Specific**: "Introduction to Neural Networks for Beginners" works better than just "AI"
2. **Optimal Slide Count**: 5-8 slides work best for most topics
3. **Review Content**: Always review and customize the generated content
4. **Replace Placeholders**: Add real images where placeholders are shown
5. **Edit as Needed**: Treat generated PPTs as a starting point

## 🐛 Troubleshooting

### Issue: "OpenAI API key not found"
- Make sure you created a `.env` file (not `.env.example`)
- Check that your API key is correctly pasted
- No quotes needed around the key

### Issue: "Template not found"
- Run `python create_templates.py` first
- Check that `templates/` folder exists with `.pptx` files

### Issue: "Rate limit exceeded"
- Your OpenAI API has usage limits
- Wait a few moments and try again
- Check your API quota at https://platform.openai.com/usage

### Issue: Generation takes too long
- Normal time is 30-60 seconds
- More slides = longer generation time
- Complex topics may take longer

## 📚 Dependencies

- `openai==1.12.0` - OpenAI API integration
- `python-pptx==0.6.23` - PowerPoint creation
- `streamlit==1.31.0` - Web interface
- `python-dotenv==1.0.1` - Environment variable management

## 🎓 Learning Notes

This project is designed to be beginner-friendly:
- **Function-based**: No complex classes
- **Well-commented**: Every step explained
- **Simple structure**: Easy to navigate
- **Clear flow**: Follow the code from start to finish

### Code Reading Order (for learning):
1. Start with `create_templates.py` to understand template creation
2. Read `ppt_generator.py` to see the core logic
3. Check `app.py` to see how the UI connects to the logic

## 🔐 Security Notes

- Never commit your `.env` file to version control
- Keep your OpenAI API key private
- The `.env.example` file is safe to share (it doesn't contain the actual key)

## 📄 License

This project is open for educational purposes. Feel free to modify and use for learning!

## 🤝 Contributing

Want to add features? Ideas for improvement:
- Add more templates
- Support for images from URLs
- Export to Google Slides format
- Batch generation
- Custom color schemes
- Chart and table generation

## 🆘 Getting Help

If you encounter issues:
1. Check the Troubleshooting section above
2. Review error messages carefully
3. Verify all setup steps were completed
4. Check OpenAI API status

## 🎉 Acknowledgments

- Built with [OpenAI GPT](https://openai.com/)
- UI powered by [Streamlit](https://streamlit.io/)
- PPT generation using [python-pptx](https://python-pptx.readthedocs.io/)

---

**Happy Presenting! 📊✨**
