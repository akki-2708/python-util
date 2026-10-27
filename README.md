#This is python utility.

python -m venv <env_name>

cd <env_name>\Scripts

activate.bat

# for pdf generation
py -m pip install openpyxl pillow reportlab

# For adding watermark on images.
py watermark.py

# To generate invoices.
py generate_invoice.py --excel orders/sample_order.xlsx

# To generate invoices with doscount.
py generate_invoice.py --excel orders/sample_order.xlsx --discount 5%


# To generate invoices with fixed shipping charges (added after discount).
py generate_invoice.py --excel orders/sample_order.xlsx --shipping 100
py generate_invoice.py --excel orders/sample_order.xlsx --discount 5% --shipping 100
