import subprocess
import sys
import os

def clear_screen():
    """Clear the terminal screen based on the operating system."""
    os.system('cls' if os.name == 'nt' else 'clear')

def print_menu():
    """Print the main menu options."""
    print("\n=== Finance Program Runner ===")
    print("1. Run bank_feeds.py")
    print("2. Run collate_spreadsheets.py")
    print("3. Run BudgetUpdater.py")
    print("4. Run all programs in sequence")
    print("5. Exit")
    print("\nPlease enter your choice (1-5): ")

def run_script(script_name):
    """Run a Python script and wait for it to complete."""
    try:
        print(f"\nRunning {script_name}...")
        # Use the same Python interpreter that's running this script
        python_executable = sys.executable
        result = subprocess.run([python_executable, script_name], check=True)
        print(f"\n{script_name} completed successfully!")
        return True
    except subprocess.CalledProcessError as e:
        print(f"\nError running {script_name}. Error code: {e.returncode}")
        return False
    except FileNotFoundError:
        print(f"\nError: {script_name} not found in the current directory.")
        return False
    except Exception as e:
        print(f"\nUnexpected error running {script_name}: {str(e)}")
        return False

def wait_for_user():
    """Wait for user input before continuing."""
    input("\nPress Enter to continue to the next program...")

def run_all_programs():
    """Run all programs in sequence with user confirmation between each."""
    programs = ['bank_feeds.py', 'collate_spreadsheets.py', 'BudgetUpdater.py']
    
    for i, program in enumerate(programs, 1):
        if not run_script(program):
            print("\nStopping sequence due to error.")
            return
        
        # Don't wait for user input after the last program
        if i < len(programs):
            wait_for_user()

def main():
    while True:
        clear_screen()
        print_menu()
        
        try:
            choice = input().strip()
            
            if choice == '1':
                run_script('bank_feeds.py')
                wait_for_user()
            elif choice == '2':
                run_script('collate_spreadsheets.py')
                wait_for_user()
            elif choice == '3':
                run_script('BudgetUpdater.py')
                wait_for_user()
            elif choice == '4':
                run_all_programs()
                wait_for_user()
            elif choice == '5':
                print("\nExiting program. Goodbye!")
                break
            else:
                print("\nInvalid choice. Please enter a number between 1 and 5.")
                wait_for_user()
                
        except KeyboardInterrupt:
            print("\n\nProgram interrupted by user. Exiting...")
            break
        except Exception as e:
            print(f"\nAn unexpected error occurred: {str(e)}")
            wait_for_user()

if __name__ == "__main__":
    main() 