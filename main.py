import googleapiclient.discovery
import pandas as pd
import os

# Load API key from environment variable or directly provide it here (not recommended for production).
DEVELOPER_KEY = "AIzaSyC-4LtqR9RtPKFqG1_LuqUMBgPlETQ6Q20"

# API service details
api_service_name = "youtube"
api_version = "v3"

# Initialize the YouTube API client
youtube = googleapiclient.discovery.build(api_service_name, api_version, developerKey=DEVELOPER_KEY)

# Function to fetch channel uploads playlist ID
def get_uploads_playlist(channel_url):
    handle = channel_url.split('@')[-1]  # Extract the handle from the URL
    channel_request = youtube.search().list(
        part="snippet",
        q=f"@{handle}",
        type="channel",
        maxResults=1
    )
    channel_response = channel_request.execute()
    if channel_response["items"]:
        channel_id = channel_response["items"][0]["id"]["channelId"]
        channel_details_request = youtube.channels().list(
            part="contentDetails",
            id=channel_id
        )
        channel_details_response = channel_details_request.execute()
        uploads_playlist_id = channel_details_response["items"][0]["contentDetails"]["relatedPlaylists"]["uploads"]
        return uploads_playlist_id
    else:
        raise ValueError("Channel handle not found or invalid.")

# Function to fetch videos from a playlist
def fetch_videos(playlist_id):
    videos = []
    next_page_token = None
    while True:
        playlist_request = youtube.playlistItems().list(
            part="snippet,contentDetails",
            playlistId=playlist_id,
            maxResults=50,
            pageToken=next_page_token
        )
        playlist_response = playlist_request.execute()
        for item in playlist_response["items"]:
            videos.append(item["contentDetails"]["videoId"])
        next_page_token = playlist_response.get("nextPageToken")
        if not next_page_token:
            break
    return videos

# Function to fetch video details
def fetch_video_details(video_ids):
    video_data = []
    for i in range(0, len(video_ids), 50):
        video_request = youtube.videos().list(
            part="snippet,statistics,contentDetails",
            id=",".join(video_ids[i:i+50])
        )
        video_response = video_request.execute()
        for item in video_response["items"]:
            snippet = item["snippet"]
            statistics = item.get("statistics", {})
            content_details = item["contentDetails"]
            video_data.append([
                item["id"],
                snippet["title"],
                snippet.get("description", ""),
                snippet["publishedAt"],
                statistics.get("viewCount", 0),
                statistics.get("likeCount", 0),
                statistics.get("commentCount", 0),
                content_details["duration"],
                snippet["thumbnails"]["high"]["url"]
            ])
    return video_data

# Function to fetch comments for a video
def fetch_comments(video_id):
    comments = []
    next_page_token = None
    while len(comments) < 100:  # Limit to the latest 100 comments
        try:
            comment_request = youtube.commentThreads().list(
                part="snippet,replies",
                videoId=video_id,
                maxResults=50,
                pageToken=next_page_token
            )
            comment_response = comment_request.execute()
            for item in comment_response["items"]:
                top_comment = item["snippet"]["topLevelComment"]["snippet"]
                comments.append([
                    video_id,
                    item["snippet"]["topLevelComment"]["id"],
                    top_comment["textOriginal"],
                    top_comment["authorDisplayName"],
                    top_comment["publishedAt"],
                    top_comment.get("likeCount", 0),
                    None  # Top-level comments have no "reply to"
                ])
                # Fetch replies if available
                replies = item.get("replies", {}).get("comments", [])
                for reply in replies:
                    reply_snippet = reply["snippet"]
                    comments.append([
                        video_id,
                        reply["id"],
                        reply_snippet["textOriginal"],
                        reply_snippet["authorDisplayName"],
                        reply_snippet["publishedAt"],
                        reply_snippet.get("likeCount", 0),
                        item["snippet"]["topLevelComment"]["id"]  # Reference to the parent comment
                    ])
            next_page_token = comment_response.get("nextPageToken")
            if not next_page_token:
                break
        except Exception as e:
            print(f"Error fetching comments for video {video_id}: {e}")
            break
    return comments

# Main script
def main(channel_url):
    try:
        uploads_playlist_id = get_uploads_playlist(channel_url)
        video_ids = fetch_videos(uploads_playlist_id)
        video_data = fetch_video_details(video_ids)

        # Fetch comments for each video (latest 100)
        comments_data = []
        for video_id in video_ids:
            comments = fetch_comments(video_id)
            comments_data.extend(comments)

        # Create Excel file
        with pd.ExcelWriter("YouTube_Channel_Data.xlsx") as writer:
            video_df = pd.DataFrame(video_data, columns=[
                "Video ID", "Title", "Description", "Published Date",
                "View Count", "Like Count", "Comment Count", "Duration", "Thumbnail URL"
            ])
            comments_df = pd.DataFrame(comments_data, columns=[
                "Video ID", "Comment ID", "Comment Text", "Author Name",
                "Published Date", "Like Count", "Reply To"
            ])
            video_df.to_excel(writer, sheet_name="Video Data", index=False)
            comments_df.to_excel(writer, sheet_name="Comments Data", index=False)

        print("Data successfully saved to 'YouTube_Channel_Data.xlsx'")
    except Exception as e:
        print(f"An error occurred: {e}")

# Replace with the YouTube channel URL containing the handle
channel_url = "https://www.youtube.com/@channelhandle"
main(channel_url)
