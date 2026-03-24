import os
import json
import pandas as pd

def compile_to_csv(json_dir, csv_path):
    all_data = []
    
    for filename in os.listdir(json_dir):
        if filename.endswith(".json"):
            with open(os.path.join(json_dir, filename), "r") as f:
                file_data = json.load(f)
                file_data["source_file"] = filename # Keep track of which PDF it came from
                all_data.append(file_data)
    
    df = pd.DataFrame(all_data)
    df.to_csv(csv_path, index=False)
    print(f"Success! CSV saved to {csv_path}")

if __name__ == "__main__":
    compile_to_csv("data/output", "data/output/final_indicators.csv")