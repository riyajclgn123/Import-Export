import os
import subprocess
import shutil
import sys


def install_requirements():
    """Ensure all dependencies from requirements.txt are installed."""
    if os.path.exists("requirements.txt"):
        subprocess.run([sys.executable, "-m", "pip", "install", "-r", "requirements.txt", "--quiet"])

def run_script(script_name, *args):
    """
    Runs a Python script using subprocess with optional arguments.
    """

    python_executable = shutil.which("python")  # Finds Python in system PATH

    #python_executable = os.path.join(os.getenv('VIRTUAL_ENV'), 'Scripts', 'python.exe')
    try:
        subprocess.run([python_executable, script_name] + list(args), check=True)
    except subprocess.CalledProcessError as e:
        print(f"An error occurred while executing {script_name}: {e}")

def main():
    install_requirements()
    print("Choose an option:")
    print("1. Create Backup For Safety)")
    print("2. Import Data (Excel to Firestore)")
    print("3. Export Data (Firestore to Excel)")
    print("4. Import Backup File")


    #Add an option for the backup of existing

    choice = input("Enter the number corresponding to your choice: ")

    #In the import section, before importing the new file to the firestore, backup the original file before importing
    # append timestamp to the backup file
    if choice == '1':
        print("************************ Creating the backup file********************")
        run_script(os.path.abspath("Export.py"), '--backup')
    elif choice == '2':
        print("************************ Importing the backup data ********************")
        run_script(os.path.abspath("Import.py"))
    elif choice == '3':
        print("************************ Exporting the Data ********************")
        run_script(os.path.abspath("Export.py"))
    elif choice == '4':
        print("************************ Importing backup file ********************")
        backup_file = "backup/" + input("Please enter the path to the backup file: ")
        if os.path.exists(backup_file):
            run_script(os.path.abspath("Import.py"), backup_file)
        else:
            print("The specified backup file does not exist. Please check the path and try again.")
    else:
        print("Invalid choice. Please enter 1 or 2.")

if __name__ == "__main__":
    main()
