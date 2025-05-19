import os
import subprocess
import argparse
import sys
from pathlib import Path
import logging
from typing import List, Optional
import shutil
import tkinter as tk
from tkinter import filedialog
import re # <--- ADDED THIS IMPORT

# Setup basic logging for the wrapper script
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# run_script function remains the same as in the previous version
def run_script(script_path: Path, args: List[str] = [], cwd: Optional[Path] = None, script_name: str = "") -> bool:
    """Runs a Python script using subprocess, handling potential errors."""
    if not script_path.is_file():
        logging.error(f"Script not found: {script_path}")
        return False

    command = [sys.executable, str(script_path)] + args
    script_name = script_name or script_path.name
    logging.info(f"Running {script_name}...")
    logging.info(f"  Command: {' '.join(command)}")
    if cwd:
        # Ensure CWD exists before running
        if not Path(cwd).is_dir():
            logging.error(f"Working directory not found for {script_name}: {cwd}")
            return False
        # The 'cwd' here is the working directory for the SUBPROCESS being run,
        # not for the file dialog. This is logged below.
        logging.info(f"  Working Directory for {script_name}: {cwd}")


    try:
        result = subprocess.run(
            command,
            capture_output=True,
            text=True,
            check=True,
            cwd=cwd,
            encoding='utf-8',
            errors='replace'
        )
        stdout_content = result.stdout.strip() if result.stdout else ""
        stderr_content = result.stderr.strip() if result.stderr else ""

        if stdout_content:
            logging.info(f"{script_name} Output:\n--- START ---\n{stdout_content}\n--- END ---")
        else:
            logging.info(f"{script_name} produced no standard output.")

        if stderr_content:
            logging.warning(f"{script_name} Stderr:\n--- START ---\n{stderr_content}\n--- END ---")

        logging.info(f"{script_name} completed successfully.")
        return True

    except subprocess.CalledProcessError as e:
        logging.error(f"Error running {script_name} (Return Code: {e.returncode}):")
        stdout_content = e.stdout.strip() if e.stdout else "No stdout captured."
        stderr_content = e.stderr.strip() if e.stderr else "No stderr captured."
        logging.error(f"Stdout:\n{stdout_content}")
        logging.error(f"Stderr:\n{stderr_content}")
        return False
    except FileNotFoundError:
        logging.error(f"Error: Could not find executable '{sys.executable}' or script '{script_path}' during execution.")
        return False
    except Exception as e:
        logging.error(f"An unexpected error occurred while trying to run {script_name}: {e}")
        return False


def select_excel_file() -> Optional[Path]:
    """Opens a file dialog for the user to select an Excel file.
    The dialog will initially open in the directory of this script.
    """
    root = tk.Tk()
    root.withdraw() # Hide the main tkinter window

    # Set the initial directory to where the script is located
    script_dir = Path(__file__).resolve().parent
    logging.info(f"Opening file dialog with initial directory: {script_dir}")

    file_path = filedialog.askopenfilename(
        title="Select Input Excel File",
        initialdir=str(script_dir), # Set the initial directory
        filetypes=(("Excel files", "*.xlsx *.xls"), ("All files", "*.*"))
    )
    root.destroy() # Destroy the root window after selection

    if file_path:
        logging.info(f"File selected: {file_path}")
        return Path(file_path)
    else:
        logging.info("File selection cancelled.")
        return None

def main():
    parser = argparse.ArgumentParser(
        description="Automate JSON creation and Invoice generation from an input Excel file."
    )
    parser.add_argument(
        "-i", "--input",
        help="Path to the input Excel file (e.g., 'input.xlsx'). If not provided, a file dialog will open.",
        type=str,
        default=None
    )
    parser.add_argument(
        "--fob",
        action="store_true",
        help="Only generate the FOB version of the invoice." # Clarified help
    )
    parser.add_argument(
        "--custom",
        action="store_true",
        help="Only generate the CUSTOM version of the invoice." # Clarified help
    )

    args = parser.parse_args()

    input_excel_path_str = args.input
    input_excel_path: Optional[Path] = None

    if input_excel_path_str:
        input_excel_path = Path(input_excel_path_str).resolve()
        logging.info(f"Input Excel file provided via command line: {input_excel_path}")
    else:
        logging.info("Input Excel file not provided via command line. Opening file dialog...")
        selected_path = select_excel_file() # This function now sets initialdir
        if selected_path:
            input_excel_path = selected_path.resolve()
            # logging.info(f"Input Excel file selected via dialog: {input_excel_path}") # Logged in select_excel_file
        else:
            # logging.info("No file selected from dialog. Exiting.") # Logged in select_excel_file
            sys.exit(0)

    if not input_excel_path or not input_excel_path.is_file():
        logging.error(f"Input Excel file not found or invalid: {input_excel_path}")
        sys.exit(1)

    # Get the directory of the input excel file, which will become the base for the main output_dir
    base_output_dir = input_excel_path.parent
    # Create output directory within the input file's directory, named after the input file (without extension)
    output_dir_name = input_excel_path.stem
    output_dir = base_output_dir / output_dir_name
    output_dir.mkdir(parents=True, exist_ok=True)
    logging.info(f"Main output directory set to: {output_dir}")


    # Define all required sub-directories relative to the main output directory
    json_output_dir = output_dir / "json_output"
    invoice_output_dir = output_dir / "invoice_output"

    # Create required sub-directories
    json_output_dir.mkdir(parents=True, exist_ok=True)
    invoice_output_dir.mkdir(parents=True, exist_ok=True)
    logging.info(f"JSON output will be in: {json_output_dir}")
    logging.info(f"Invoice output will be in: {invoice_output_dir}")


    # --- Define Project Structure & Validate Paths ---
    # project_root is the directory containing this wrapper script
    project_root = Path(__file__).resolve().parent
    create_json_dir = project_root / "create_json"
    invoice_gen_dir = project_root / "invoice_gen"
    create_json_script = create_json_dir / "main.py"
    invoice_gen_script = invoice_gen_dir / "generate_invoice.py"
    template_dir = project_root / "invoice_gen" / "TEMPLATE"
    config_dir = project_root / "invoice_gen" / "config"

    identifier = input_excel_path.stem
    
    # --- MODIFIED PREFIX EXTRACTION ---
    match = re.match(r'([A-Za-z]+)', identifier)
    prefix = match.group(1) if match else ''
    # --- END OF MODIFICATION ---

    if not prefix:
        logging.error(f"Could not extract alphabetic prefix from filename: {identifier}") # Updated error message slightly
        sys.exit(1)

    # Validate paths that must exist beforehand
    essential_paths_to_check = {
        "JSON creation script": create_json_script,
        "Invoice generation script": invoice_gen_script,
        "Template directory": template_dir,
        "Configuration directory": config_dir,
    }
    for name, path_to_check in essential_paths_to_check.items():
        if (path_to_check.is_file() if "script" in name else path_to_check.is_dir()):
            logging.info(f"Found {name}: {path_to_check}")
        else:
            logging.error(f"{name} not found at expected location: {path_to_check}")
            if name == "Configuration directory":
                logging.error("Please ensure the directory exists and is correct.")
            sys.exit(1)


    # --- Step 1: Run create_json/main.py ---
    create_json_args = [
        "--input-excel", str(input_excel_path),
        "--output-dir", str(json_output_dir)
    ]
    logging.info(f"Running JSON creation step (create_json/main.py) using input: {input_excel_path}")
    # The cwd for create_json_script is create_json_dir, allowing it to find relative files if needed
    if not run_script(create_json_script, args=create_json_args, cwd=create_json_dir, script_name="create_json"):
        logging.error("JSON creation script failed. Aborting.")
        sys.exit(1)

    # --- Step 2: Verify JSON Output ---
    expected_json_path = json_output_dir / f"{identifier}.json"
    if not expected_json_path.is_file():
        logging.error(f"Expected JSON output file was not found: {expected_json_path}")
        logging.error("Check the output/logs of the create_json script for errors.")
        sys.exit(1)
    logging.info(f"JSON file successfully created: {expected_json_path}")

    # --- Step 3: Verify Expected Config for invoice_gen ---
    expected_config_path = config_dir / f"{prefix}_config.json"
    logging.info(f"Invoice generation step will expect main config file: {expected_config_path}")
    if not expected_config_path.is_file():
        logging.error(f"Expected main config file '{expected_config_path}' not found in '{config_dir}'.")
        logging.error(f"Please ensure the required config file for prefix '{prefix}' exists (e.g., {prefix}_config.json).") # Dynamic prefix in error
        sys.exit(1)

    # --- Step 4: Run invoice_gen/generate_invoice.py for each mode ---
    active_modes = []
    # If specific mode flags are given, only run those modes.
    # Otherwise, run all defined modes.
    if args.fob:
        active_modes.append(("fob", ["--fob"]))
    if args.custom:
        active_modes.append(("custom", ["--custom"]))

    if not active_modes: # If no specific mode flag was given, run all standard modes
        active_modes = [
            ("normal", []),
            ("fob", ["--fob"]),
            ("custom", ["--custom"])
        ]

    all_successful_invoice_generations = True
    generated_files_info = []

    for mode_name, mode_flags in active_modes:
        logging.info(f"--- Processing {mode_name.upper()} mode for invoice generation ---")

        output_filename = f"CT&INV&PL {identifier} {mode_name.upper()}.xlsx"
        invoice_gen_args = [
            str(expected_json_path), # First argument is the path to the generated JSON
            "--output", str(invoice_output_dir / output_filename),
            "--templatedir", str(template_dir),
            "--configdir", str(config_dir), # Main config dir (for JF_config.json etc.)
        ] + mode_flags # Add specific flags for this mode (e.g., --fob, --custom)

        logging.info(f"Running Invoice generation (invoice_gen/generate_invoice.py) to create: {output_filename}")
        # The cwd for invoice_gen_script is invoice_gen_dir
        if not run_script(invoice_gen_script, args=invoice_gen_args, cwd=invoice_gen_dir, script_name=f"invoice_gen ({mode_name})"):
            logging.error(f"Invoice generation script failed for {mode_name} mode.")
            all_successful_invoice_generations = False
        else:
            generated_files_info.append(f"{len(generated_files_info) + 1}. {mode_name.capitalize()}: {output_filename}")

    # --- Final Summary ---
    if generated_files_info: # If at least one file was attempted/logged
        if all_successful_invoice_generations and len(generated_files_info) == len(active_modes):
            logging.info("--- Automation Completed Successfully ---")
        elif generated_files_info:
             logging.warning("--- Automation Completed with some errors or not all modes run/successful ---")
        else: # Should not happen if generated_files_info is populated but all_successful is false
             logging.warning("--- Automation Completed, but with issues noted above. ---")

        logging.info(f"All outputs are in directory: {output_dir.resolve()}")
        if generated_files_info:
            logging.info("Generated invoice versions:")
            for line in generated_files_info:
                logging.info(line)
        else:
            logging.info("No invoice versions were generated successfully.")

    elif not active_modes:
        logging.warning("--- Automation SKIPPED --- No invoice generation modes were specified or active.")
    else:
        logging.error("--- Automation FAILED --- No invoice files were generated. Review logs for errors.")

    if not all_successful_invoice_generations and generated_files_info:
        sys.exit(2) # Exit with a different code if some invoices failed
    elif not generated_files_info and active_modes: # Attempted but none succeeded
        sys.exit(3) # Exit with a different code if no invoices generated despite attempts


if __name__ == "__main__":
    main()