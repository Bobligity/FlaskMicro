import json

data = {
  "operation": "request_fields",
  "write_ts": "Hello World"
}

with open("task.json", "w") as write_file:
    json.dump(data, write_file)