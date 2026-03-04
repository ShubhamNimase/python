import pandas as pd

def create_sample_excel():
    data = {
        "Name": ["John Doe", "Jane Smith", "Alice Johnson"],
        "Email ID": ["john.doe@example.com", "jane.smith@example.com", "alice.johnson@example.com"]
    }

    df = pd.DataFrame(data)
    df.to_excel("contacts.xlsx", index=False)
    print("Sample 'contacts.xlsx' created successfully.")

if __name__ == "__main__":
    create_sample_excel()
