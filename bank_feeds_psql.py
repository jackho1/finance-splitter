import requests
from datetime import datetime, timedelta
import psycopg2
from psycopg2 import sql
from config import (  # Import settings from config.py
    POCKETSMITH_API_KEY,
    ULTIMATE_AWARDS_CC_ID,
    PEOPLE,
    DB_CONFIG,
    DAYS_TO_FETCH
)

# Function to fetch bank feed transactions from PocketSmith API
def fetch_transactions(start_date=None):
    page = 1
    all_transactions = []

    while True:
        url = f"https://api.pocketsmith.com/v2/accounts/{ULTIMATE_AWARDS_CC_ID}/transactions?page={page}&start_date={start_date}&end_date={datetime.now()}"
        
        headers = {"accept": "application/json", "X-Developer-Key": POCKETSMITH_API_KEY}

        print(f"Fetching transactions from PocketSmith API - Page {page}...")
        response = requests.get(url, headers=headers)

        if response.status_code != 200:
            return all_transactions

        transactions = response.json()

        if not transactions:  # If there are no transactions, we stop fetching
            break
        
        # Parse transactions to match our format
        for tx in transactions:
            if tx.get('category') and 'title' in tx['category']:
                category_title = tx['category']['title']
            else:
                category_title = ""
            all_transactions.append({
                'id': tx['id'],
                'date': tx['date'],
                'description': tx['payee'],
                'bank_category': category_title,
                'amount': tx['amount']
            })

        page += 1  # Move to the next page for the next iteration
    return all_transactions

# Function to auto-label bank categories
def auto_label_bank_category(bank_category):
    """
    Automatically label transactions based on bank category.
    
    Returns the appropriate person label from PEOPLE list or None for shared/unlabeled categories.
    """
    no_label_categories = ["Dining", "Travel"]
    first_person_categories = ["Personal Items", "Personal Care", "Hobbies", 
                              "Entertainment/Recreation", "Vehicle", "Gym", "Fuel"]
    
    if not bank_category or bank_category in no_label_categories:
        return None
    elif bank_category in first_person_categories:
        return PEOPLE[0]
    else:
        return "Both"

# Function to categorize and label transactions
def categorize_and_label_transactions(transactions):
    categorized_transactions = []
    
    for tx in transactions:
        description = tx['description']
        amount = tx['amount']
        bank_category = tx['bank_category']

        label = auto_label_bank_category(bank_category)
        category = None
        label = label if label else None

        categorized_transactions.append({
            'id': tx['id'],
            'date': tx['date'],
            'description': description,
            'amount': amount,
            'category': category,
            'bank_category': bank_category,
            'label': label
        })
    
    return categorized_transactions

# Function to insert transactions into PostgreSQL database
def insert_transactions(transactions):
    try:
        conn = psycopg2.connect(**DB_CONFIG)
        cursor = conn.cursor()

        for tx in transactions:
            insert_query = sql.SQL("""
                INSERT INTO shared_transactions (id, date, description, amount, category, bank_category, label)
                VALUES (%s, %s, %s, %s, %s, %s, %s)
                ON CONFLICT (id) DO NOTHING
            """)
            cursor.execute(insert_query, (
                tx['id'],
                tx['date'],
                tx['description'],
                tx['amount'],
                tx['category'],
                tx['bank_category'],
                tx['label']
            ))

        conn.commit()
        print("Transactions inserted successfully.")

    except Exception as e:
        print(f"Error inserting transactions: {e}")
    finally:
        cursor.close()
        conn.close()

# Main function to fetch, categorize, and save transactions
def main():
    start_date = (datetime.now() - timedelta(days=DAYS_TO_FETCH)).strftime('%Y-%m-%d')
    print(f"Fetching transactions starting from {start_date} to {datetime.now().strftime('%Y-%m-%d')} (spans {DAYS_TO_FETCH} days).")

    transactions = fetch_transactions(start_date)
    categorized_transactions = categorize_and_label_transactions(transactions)
    insert_transactions(categorized_transactions)

if __name__ == "__main__":
    main()