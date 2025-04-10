import pandas as pd
from pathlib import Path
from fpdf import FPDF
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
import os
from dotenv import load_dotenv
import logging
import warnings

# Suppress warnings
warnings.filterwarnings('ignore')

# Load environment variables
load_dotenv()

# Constants
PAYSLIPS_DIR = Path("payslips")
PAYSLIPS_DIR.mkdir(exist_ok=True)

# Logging configuration
LOG_FILE = "payslip_generator.log"
logging.basicConfig(filename=LOG_FILE, level=logging.ERROR,
                    format='%(asctime)s - %(levelname)s - %(message)s')

# Email configuration (using environment variables)
SMTP_SERVER = os.getenv('SMTP_SERVER', 'smtp.gmail.com')
SMTP_PORT = int(os.getenv('SMTP_PORT', 587))
SENDER_EMAIL = os.getenv('SENDER_EMAIL')
SENDER_PASSWORD = os.getenv('SENDER_PASSWORD')


class PayslipGenerator:
    def __init__(self, excel_file: str, currency_symbol: str = "$") -> None:
        self.excel_file = Path(excel_file)
        self.employee_data: pd.DataFrame | None = None
        self.currency_symbol = currency_symbol
        self._load_initial_data()  # Load data during initialization

    def _load_initial_data(self) -> bool:
        """Load employee data from Excel file and attempt to fix column name issues."""
        try:
            self.employee_data = pd.read_excel(self.excel_file, header=0)
            print("Employee Data loaded successfully:")
            print(self.employee_data)
            print("\nData Summary:")
            print(self.employee_data.info())

            # --- Attempt to fix 'Allowances' column name ---
            expected_allowances = 'Allowances'
            found_allowances = None
            actual_columns = list(self.employee_data.columns)
            print("--- DEBUG: Actual Column Names ---")
            print(actual_columns)
            print("--- END DEBUG ---")

            for col in actual_columns:
                if col.lower().strip() == expected_allowances.lower():
                    found_allowances = col
                    break

            if found_allowances and found_allowances != expected_allowances:
                print(f"DEBUG: Renaming column '{found_allowances}' to '{expected_allowances}'")
                self.employee_data.rename(columns={found_allowances: expected_allowances}, inplace=True)
                print("DEBUG: Column names after potential rename:", list(self.employee_data.columns))

            if not self._validate_data():
                return False
            return True
        except FileNotFoundError:
            error_message = f"Error: Excel file not found at {self.excel_file}"
            logging.error(error_message)
            print(error_message)
            return False
        except Exception as e:
            error_message = f"Error loading employee data: {e}"
            logging.error(error_message)
            print(error_message)
            return False

    def _validate_data(self) -> bool:
        """Validates the employee data."""
        required_columns = {'Employee ID', 'Name', 'Email', 'Basic Salary', 'Allowances', 'Deductions'}
        if not required_columns.issubset(self.employee_data.columns):
            missing_columns = required_columns - set(self.employee_data.columns)
            logging.error(f"Missing required columns: {missing_columns}")
            print(f"Error: Missing required columns in Excel file: {missing_columns}")
            return False
        return True

    def calculate_net_salary(self, row: pd.Series) -> float:
        """Calculate net salary for an employee."""
        return row['Basic Salary'] + row['Allowances'] - row['Deductions']

    def generate_payslip(self, employee: pd.Series) -> Path:
        """Generate PDF payslip for an employee."""

        pdf = FPDF()
        pdf.add_page()

        # Set font and styles
        pdf.set_font('Arial', 'B', 16)
        pdf.cell(0, 10, 'Uncommon.org Payslip', 0, 1, 'C')
        pdf.ln(10)

        # Employee Details
        pdf.set_font('Arial', '', 12)
        pdf.cell(0, 10, f"Employee ID: {employee['Employee ID']}", 0, 1)
        pdf.cell(0, 10, f"Name: {employee['Name']}", 0, 1)
        pdf.ln(5)

        # Salary Details
        pdf.set_font('Arial', 'B', 12)
        pdf.cell(0, 10, 'Salary Details', 0, 1)
        pdf.set_font('Arial', '', 12)

        pdf.cell(90, 10, 'Basic Salary:', 1)
        pdf.cell(0, 10, f"{self.currency_symbol} {employee['Basic Salary']:.2f}", 1, 1)

        pdf.cell(90, 10, 'Allowances:', 1)
        pdf.cell(0, 10, f"{self.currency_symbol} {employee['Allowances']:.2f}", 1, 1)

        pdf.cell(90, 10, 'Deductions:', 1)
        pdf.cell(0, 10, f"{self.currency_symbol} {employee['Deductions']:.2f}", 1, 1)

        # Net Salary
        pdf.set_font('Arial', 'B', 12)
        pdf.cell(90, 10, 'Net Salary:', 1)
        net_salary = self.calculate_net_salary(employee)
        pdf.cell(0, 10, f"{self.currency_symbol} {net_salary:.2f}", 1, 1)

        # Save the PDF
        filename = PAYSLIPS_DIR / f"{employee['Employee ID']}.pdf"
        pdf.output(str(filename))
        return filename

    def send_email(self, recipient_email: str, payslip_path: Path) -> bool:
        """Send payslip via email."""
        try:
            # Create message
            msg = MIMEMultipart()
            msg['From'] = SENDER_EMAIL
            msg['To'] = recipient_email
            msg['Subject'] = "Your Payslip for This Month"

            # Email body
            body = f"""Dear Employee,

Please find attached your payslip for this month.

Best regards,
Uncommon.org HR Department
"""
            msg.attach(MIMEText(body, 'plain'))

            # Attach payslip
            with open(payslip_path, 'rb') as f:
                attach = MIMEApplication(f.read(), _subtype="pdf")
                attach.add_header('Content-Disposition', 'attachment',
                                filename=f"payslip_{payslip_path.stem}.pdf")
                msg.attach(attach)

            # Send email
            with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
                server.starttls()
                server.login(SENDER_EMAIL, SENDER_PASSWORD)
                server.send_message(msg)

            print(f"Email sent successfully to {recipient_email}")
            return True

        except smtplib.SMTPException as smtp_error:
            error_message = f"SMTP error sending email to {recipient_email}: {smtp_error}"
            logging.error(error_message)
            print(error_message)
            return False
        except Exception as e:
            error_message = f"Failed to send email to {recipient_email}: {e}"
            logging.error(error_message)
            print(error_message)
            return False

    def process_all(self) -> bool:
        """Process all employees."""
        if self.employee_data is None:
            print("Error: Employee data not loaded. Please check the Excel file.")
            return False

        success_count = 0
        total_employees = len(self.employee_data)

        for _, employee in self.employee_data.iterrows():
            try:
                print(f"\nProcessing employee: {employee['Name']} (ID: {employee['Employee ID']})")

                # Generate payslip
                payslip_path = self.generate_payslip(employee)
                print(f"Payslip generated: {payslip_path}")

                # Send email
                if self.send_email(employee['Email'], payslip_path):
                    success_count += 1

            except Exception as e:
                error_message = f"Error processing employee {employee['Name']}: {e}"
                logging.error(error_message)
                print(error_message)

        print(f"\nProcessed {success_count}/{total_employees} employees successfully")
        return success_count == total_employees


if __name__ == "__main__":
    # Create and run the payslip generator
    generator = PayslipGenerator('employees.xlsx', currency_symbol="$")

    if generator.process_all():
        print("\nPayslip generation and emailing completed successfully!")
    else:
        print("\nPayslip generation completed with some errors. Please check the output and log file.")