
import requests
import sys

task_id = sys.argv[1]
url = f"http://localhost:5000/extraction-status/{task_id}"

response = requests.get(url)
print(response.json())
