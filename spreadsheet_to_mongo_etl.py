from pymongo import MongoClient

# MongoDB connection details
MONGO_URI = "mongodb://localhost:27017/"
db_name = "spreadsheet_data"
collection_name = "final_uso"

data = pd.read_excel(file_path, engine='openpyxl')
records = data.to_dict(orient='records')

client = MongoClient(MONGO_URI)
db = client[db_name]
collection = db[collection_name]

result = collection.insert_many(records)

