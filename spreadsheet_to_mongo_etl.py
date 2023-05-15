from pymongo import MongoClient

# MongoDB connection details
MONGO_URI = "mongodb://localhost:27017/"

db_name = "spreadsheet_data"
collection_name = "final_uso"

def load_data_to_mongodb(file_path):
    try:
        data = pd.read_excel(file_path, engine='openpyxl')
        records = data.to_dict(orient='records')  # Convert DataFrame to list of dictionaries
				     
        client = MongoClient(MONGO_URI)
        db = client[db_name]
        collection = db[collection_name]

        result = collection.insert_many(records)
    
    except Exception as e:
        raise
    
    finally:
        client.close()

if __name__ == "__main__":
    file_path = "FINAL_USO.xlsx"  # Path to your Excel file
    load_data_to_mongodb(file_path)
