
import requests
import pandas as pd
import time
import os

# =================================
# 1. CONFIGURATION
# =================================

BASE_URL = "https://peertube2.cpy.re/api/v1/videos"

HEADERS = {
    "User-Agent": "OTT-Learning-Project/1.0"
}

VIDEOS_PER_PAGE = 20
MAX_PAGES = 3

# =================================
# 2. FETCH VIDEO DATA
# =================================

all_videos = []

for page in range(MAX_PAGES):

    start = page * VIDEOS_PER_PAGE

    params = {
        "start": start,
        "count": VIDEOS_PER_PAGE,
        "sort": "-publishedAt",
        "nsfw": "false"
    }

    print(f"Fetching page {page + 1}...")

    try:
        response = requests.get(
            BASE_URL,
            headers=HEADERS,
            params=params,
            timeout=30
        )

        response.raise_for_status()

        result = response.json()

        videos = result.get("data", [])

        if not videos:
            print("No more videos found.")
            break

        # =================================
        # 3. EXTRACT VIDEO INFORMATION
        # =================================

        for video in videos:

            channel = video.get("channel") or {}
            account = video.get("account") or {}
            category = video.get("category") or {}
            language = video.get("language") or {}

            video_data = {

                "Video ID": video.get("id"),

                "Video Title": video.get("name"),

                "Description": video.get(
                    "description"
                ),

                "Channel Name": channel.get(
                    "displayName"
                ),

                "Creator": account.get(
                    "displayName"
                ),

                "Category": category.get(
                    "label"
                ),

                "Language": language.get(
                    "label"
                ),

                "Duration": video.get(
                    "duration"
                ),

                "Views": video.get(
                    "views"
                ),

                "Likes": video.get(
                    "likes"
                ),

                "Published Date": video.get(
                    "publishedAt"
                ),

                "Is Live": video.get(
                    "isLive"
                )
            }

            all_videos.append(video_data)

        print("Videos collected:", len(all_videos))

        # Be polite to the server
        time.sleep(1)

    except requests.exceptions.RequestException as error:

        print("Request error:", error)
        break

    except ValueError as error:

        print("JSON parsing error:", error)
        break

# =================================
# 4. SAVE DATA TO EXCEL
# =================================

if all_videos:

    df = pd.DataFrame(all_videos)

    os.makedirs("output", exist_ok=True)

    output_file = "output/peertube_data.xlsx"

    df.to_excel(
        output_file,
        index=False,
        sheet_name="OTT Videos"
    )

    print("\nScraping completed!")
    print("Total videos:", len(df))
    print("Excel saved at:", output_file)

else:

    print("No video data collected.")
