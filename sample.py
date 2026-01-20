import requests

# url = "https://api.chatwork.com/v2/me"
url = "https://api.chatwork.com/v2/rooms"

headers = {"accept": "application/json", "x-chatworktoken": "6e1f2ca736c32a04246486309350df88"}

response = requests.get(url, headers=headers)

# 'name'キーのみ表示する
for room in response.json():
    print(room['name'])
# print(response.json())