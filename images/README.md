# Pexels Image Integration

This folder contains the Pexels API integration for automatic image downloads.

## Setup Instructions

1. **Get Your Pexels API Key:**
   - Go to https://www.pexels.com/api/
   - Click "Get Started" or "Sign Up"
   - Create a free account
   - Generate your API key from the dashboard

2. **Add API Key to .env File:**
   - Open the `.env` file in the project root
   - Add your Pexels API key:
   ```
   PEXELS_API_KEY=your_actual_api_key_here
   ```

3. **That's It!**
   - Images will now be automatically downloaded from Pexels when generating presentations
   - Downloaded images are cached in `images/downloads/` folder
   - If Pexels API fails, the system falls back to gradient placeholders

## How It Works

- When generating a PowerPoint presentation, the system analyzes each slide title
- Relevant keywords are extracted from the title
- Pexels API is called to search for matching images
- Best matching image is downloaded and inserted into the slide
- Images are cached locally to avoid re-downloading

## Features

- ✅ Automatic keyword extraction from slide titles
- ✅ Smart image search using Pexels API
- ✅ Local caching to save API calls
- ✅ Graceful fallback to placeholders if images unavailable
- ✅ Landscape-oriented images for better slide fit

## Folder Structure

```
images/
├── pexels_fetcher.py     # Main Pexels API integration code
├── downloads/            # Downloaded images cache (auto-created)
└── README.md            # This file
```

## Notes

- Pexels API is free with rate limits (200 requests/hour, 20,000/month)
- Images are for educational and personal use
- Downloaded images are cached to minimize API calls
- Old cached images are automatically cleaned up
