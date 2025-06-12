import requests
 
url = "https://source.unsplash.com/random/"
response = requests.get(url)
print(response)