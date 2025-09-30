# File paths
DEFAULT_EXCEL_PATH = r'E:\Finance1\InteractiveInvestor\II_20250917.xlsx'  # Fallback default
DEFAULT_BASE_PATH = r'E:\Finance1\InteractiveInvestor'
TRANSACTIONS_CSV_PATH = r'C:\Users\moore\Downloads\Transactions.csv'
INVESTMENTS_CSV_PATH = r'C:\Users\moore\Downloads\Investments.csv'

# Transactions constants
TRANSACTIONS_SHEET = 'Transactions'
TRANSACTIONS_FORMULA = '=IF(ISERROR(VLOOKUP(D{row},MapName!A:D,4,0)),VLOOKUP(H{row},MapEdgeCases!A:B,2,0),VLOOKUP(D{row},MapName!A:D,4,0))'
TRANSACTIONS_MAX_COLUMN = 13  # A:M

# Investments constants
INVESTMENTS_SHEET = 'Investments'
INVESTMENTS_FORMULA = '=IF(ISERROR(VLOOKUP(B{row},MapName!A:D,4,0)),"",VLOOKUP(B{row},MapName!A:D,4,0))'
INVESTMENTS_MAX_COLUMN = 1  # A only
INVESTMENTS_START_CELL = 'B2'