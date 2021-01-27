import requests

api_key = "309f61194cfeee533c6ba70546e8d74378ce5"
# the URL you want to shorten
url = "https://www.thepythoncode.com/article/make-url-shortener-in-python"
# preferred name in the URL
api_url = f"https://cutt.ly/api/api.php?key={api_key}&short={url}"
# or
# api_url = f"https://cutt.ly/api/api.php?key={api_key}&short={url}&name=some_unique_name"
# make the request
data = requests.get(api_url).json()["url"]
if data["status"] == 7:
    # OK, get shortened URL
    shortened_url = data["shortLink"]
    print("Shortened URL:", shortened_url)
else:
    print("[!] Error Shortening URL:", data)