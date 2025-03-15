# finance-splitter
Splits transactions between two people

## Requirements

- Python 3.6 or higher
- Required Python packages (install via pip):
  ```bash
  pip install openpyxl subprocess
  ```

## Configuration

To set up the configuration:

1. Copy the template file:
   ```bash
   cp config_template.py config.py
   ```
2. Edit `config.py` with your actual configuration values.

## Running the Application

This application consists of several scripts that work together to manage and split financial transactions.

### Using the Program Runner

The easiest way to run the application is to use the included program runner:

1. Open your terminal or command prompt
2. Navigate to the project directory
3. Run the program runner:
   ```bash
   python3 run_programs.py
   ```
4. Select from the menu options:
   - Option 1: Run bank_feeds.py (fetches transactions from your bank)
   - Option 2: Run collate_spreadsheets.py (combines transaction data)
   - Option 3: Run BudgetUpdater.py (updates your budget)
   - Option 4: Run all programs in sequence
   - Option 5: Exit

### Running Scripts Individually

You can also run each script individually if needed:

```bash
python3 bank_feeds.py
python3 collate_spreadsheets.py
python3 BudgetUpdater.py
```