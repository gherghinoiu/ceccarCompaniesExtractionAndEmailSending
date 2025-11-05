
import requests

url = "http://localhost:5000/start-extraction"
data = {
    "member_region": "1", # Arad
    "region_name": "Arad"
}

response = requests.post(url, data=data)
print(response.json())
