"""
Pexels Image Fetcher
Downloads relevant images from Pexels API for presentation slides
"""

import os
import requests
from dotenv import load_dotenv

# Load environment variables
load_dotenv()


def fetch_image_from_pexels(search_query, save_path, orientation="landscape"):
    """
    Fetches a relevant image from Pexels API based on search query

    
    Args:
        search_query: Topic/keyword to search for
        save_path: Full path where image should be saved
        orientation: Image orientation ("landscape", "portrait", or "square")

    Returns:
        Path to downloaded image if successful, None otherwise
    """

    # Get API key from environment
    api_key = os.getenv('PEXELS_API_KEY')

    if not api_key:
        print("Warning: PEXELS_API_KEY not found in .env file")
        return None

    # Pexels API endpoint
    url = "https://api.pexels.com/v1/search"

    # API headers
    headers = {
        "Authorization": api_key
    }

    # Search parameters
    params = {
        "query": search_query,
        "per_page": 1,  # Get only 1 image
        "orientation": orientation,
        "size": "medium"  # Get medium-sized images
    }

    try:
        # Make API request
        print(f"Searching Pexels for: {search_query}")
        response = requests.get(url, headers=headers, params=params, timeout=10)
        response.raise_for_status()

        data = response.json()

        # Check if we got results
        if data.get('photos') and len(data['photos']) > 0:
            photo = data['photos'][0]
            image_url = photo['src']['large']  # Get large version

            # Download the image
            print(f"Downloading image from: {image_url}")
            img_response = requests.get(image_url, timeout=15)
            img_response.raise_for_status()

            # Save the image
            with open(save_path, 'wb') as f:
                f.write(img_response.content)

            print(f"Image saved: {save_path}")
            return save_path
        else:
            print(f"No images found for: {search_query}")
            return None

    except requests.exceptions.RequestException as e:
        print(f"Error fetching image: {e}")
        return None
    except Exception as e:
        print(f"Unexpected error: {e}")
        return None


def get_image_for_slide(slide_title, slide_index, output_dir="images/downloads"):
    """
    Gets a relevant image for a slide based on its title

    Args:
        slide_title: Title of the slide
        slide_index: Index of the slide (for unique filename)
        output_dir: Directory to save images

    Returns:
        Path to downloaded image or None
    """

    # Create downloads directory if it doesn't exist
    os.makedirs(output_dir, exist_ok=True)

    # Create safe filename from slide title
    safe_title = "".join(c if c.isalnum() else "_" for c in slide_title)
    safe_title = safe_title[:50]  # Limit length
    image_filename = f"slide_{slide_index}_{safe_title}.jpg"
    image_path = os.path.join(output_dir, image_filename)

    # Check if image already exists (avoid re-downloading)
    if os.path.exists(image_path):
        print(f"Using cached image: {image_path}")
        return image_path

    # Extract key search terms from slide title
    # Remove common words and keep important keywords
    search_query = extract_search_keywords(slide_title)

    # Fetch image from Pexels
    return fetch_image_from_pexels(search_query, image_path)


def extract_search_keywords(slide_title):
    """
    Extracts relevant keywords from slide title for image search

    Args:
        slide_title: Title of the slide

    Returns:
        Search query string
    """

    # Common words to remove
    stop_words = [
        "the", "a", "an", "and", "or", "but", "in", "on", "at", "to", "for",
        "of", "with", "by", "from", "as", "is", "was", "are", "were", "been",
        "be", "have", "has", "had", "do", "does", "did", "will", "would", "could",
        "should", "may", "might", "must", "can", "about", "into", "through",
        "during", "before", "after", "above", "below", "between", "under",
        "what", "why", "how", "when", "where", "who", "which", "this", "that",
        "these", "those", "introduction", "overview", "conclusion", "summary"
    ]

    # Convert to lowercase and split into words
    words = slide_title.lower().split()

    # Remove stop words and keep important keywords
    keywords = [word for word in words if word not in stop_words and len(word) > 2]

    # Join keywords back together
    if keywords:
        return " ".join(keywords[:4])  # Use first 4 keywords max
    else:
        # Fallback to original title if no keywords found
        return slide_title


def cleanup_old_images(output_dir="images/downloads", keep_recent=20):
    """
    Cleans up old downloaded images to save space

    Args:
        output_dir: Directory containing downloaded images
        keep_recent: Number of recent images to keep
    """

    try:
        if not os.path.exists(output_dir):
            return

        # Get all image files
        images = [
            os.path.join(output_dir, f)
            for f in os.listdir(output_dir)
            if f.endswith(('.jpg', '.jpeg', '.png'))
        ]

        # Sort by modification time (newest first)
        images.sort(key=lambda x: os.path.getmtime(x), reverse=True)

        # Delete older images
        if len(images) > keep_recent:
            for img_path in images[keep_recent:]:
                os.remove(img_path)
                print(f"Cleaned up old image: {img_path}")

    except Exception as e:
        print(f"Error during cleanup: {e}")


# Test function
if __name__ == "__main__":
    print("Testing Pexels Image Fetcher...")

    # Test search
    test_query = "photosynthesis plants"
    test_path = "images/downloads/test_image.jpg"

    os.makedirs("images/downloads", exist_ok=True)

    result = fetch_image_from_pexels(test_query, test_path)

    if result:
        print(f"✅ Test successful! Image saved at: {result}")
    else:
        print("❌ Test failed - check your API key in .env file")
