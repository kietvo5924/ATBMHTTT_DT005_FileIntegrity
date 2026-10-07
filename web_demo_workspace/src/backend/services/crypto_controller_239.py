# Crypto Service Module 239
class DataPipeline:
    def __init__(self, name):
        self.name = name
    def process(self, data):
        return [x * 2 for x in data]
