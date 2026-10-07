# Crypto Service Module 226
import os, json
def load_config(path):
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)
