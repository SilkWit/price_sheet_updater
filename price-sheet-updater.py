from tkinter import Tk, Label, messagebox
from tkinterdnd2 import TkinterDnD, DND_FILES
from pathlib import Path
import pandas as pd
import sqlite3

def parse_excel(file_path):
    #read the Excel file, skipping irrelevant header rows
    data = pd.read_excel(file_path, skiprows=5)  #adjust skiprows to start reading from actual data
    print("Filtered Data read from Excel:", data.head())  #debugging, filtering

    #filter to only include the necessary columns, renaming as needed - empty/null columns were being added
    data = data.rename(columns={
        "PRODUCT DESCRIPTION": "product_description",
        "Quantity": "quantity",
        "UOM": "UOM",
        "PRICE": "price"
    })
    
    #select only the relevant columns
    data = data[["product_description", "quantity", "UOM", "price"]]
    
    #fill nulls with empty strings for non-numeric fields and 0 for price
    data = data.fillna({"product_description": "", "quantity": "", "UOM": "", "price": 0.0})

    #convert dataframe rows into a list of lists for SQL insertion
    data_rows = data.values.tolist()
    print("Parsed data rows:", data_rows)  #debugging, see what's being recorded
    return data_rows

def update_or_insert_sql(data_rows):
    #connect to the sql database
    connect = sqlite3.connect('roofing_materials.db')
    cursor = connect.cursor()

    #create table if it doesn't exist
    cursor.execute('''CREATE TABLE IF NOT EXISTS RoofingMaterials (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        product_description TEXT,
                        quantity TEXT,
                        UOM TEXT,
                        price REAL)''')

    for row in data_rows:
        product_description, quantity, UOM, price = row
        
        #check if the product already exists in the database
        cursor.execute('''SELECT id, price FROM RoofingMaterials WHERE product_description = ?''', (product_description,))
        existing_product = cursor.fetchone()

        if existing_product:
            #if product exists, update the price if it has changed
            existing_price = existing_product[1]
            if existing_price != price:
                print(f"Updating price for {product_description} from {existing_price} to {price}")
                cursor.execute('''UPDATE RoofingMaterials SET price = ? WHERE product_description = ?''', (price, product_description))
        else:
            #if product doesn't exist, insert it
            print(f"Inserting new product: {product_description}")
            cursor.execute('''INSERT INTO RoofingMaterials (product_description, quantity, UOM, price)
                              VALUES (?, ?, ?, ?)''', row)

    #commit the changes and close the connection
    connect.commit()
    connect.close()
    print("Database update successful.")

def on_drop(event):
    #get the dropped file path and clean it - it wouldn't find path due to spaces and special characters in file name
    raw_file_path = event.data.strip().replace("{", "").replace("}", "")
    file_path = Path(raw_file_path)

    #handle potential path issues
    if not file_path.exists():
        file_path = Path(" ".join(raw_file_path.split()))

    #handle excel files
    if file_path.suffix.lower() == ".xlsx" and file_path.exists():
        #parse excel to get structured data
        data_rows = parse_excel(file_path)
        update_or_insert_sql(data_rows)  #insert or update data in sql
        messagebox.showinfo("File Processed", f"Content saved to roofing_materials.db")
    else:
        messagebox.showinfo("Unsupported File", "Only Excel (.xlsx) files are currently supported.")

#setting up the main window using tkinter
window = TkinterDnD.Tk()
window.title("Drag and Drop Files Here")
window.geometry("400x300")

label = Label(window, text="Drop price sheet document here", font=("Arial", 18))
label.pack(expand=True, fill="both")

#enable the window to accept dropped files
window.drop_target_register(DND_FILES)
window.dnd_bind('<<Drop>>', on_drop)

window.mainloop()
