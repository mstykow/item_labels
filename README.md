# Labels from Excel
Program to create item labels on Avery 5160 from QuickBooks Excel export file.

After an order has been placed and invoiced through QuickBooks, a company exports a memorized report containing labels for each item the customer has ordered. These labels get put onto brown bags which are then filled accordingly by staff. Often there can be hundreds of labels per day from many customers. This program creates the required labels (containing item name, invoice number, quantity, organic certification status) straight from the QuickBooks Excel export file. It is also able to create partial sheets where some of the labels have been previously used.

* See [here](https://www.avery.com/templates/5160) for the label template.

Additional comments:

* Place your QuickBooks Excel export file into the same directory as this script, run item_labels.py, and follow prompts.
* You may need to install the following modules:
  * openpyxl: pip install openpyxl
  * pylabels: pip install pylabels
  * reportlab: pip install reportlab
* Tested and created on QuickBooks Pro 2015.
