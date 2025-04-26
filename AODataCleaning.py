#%%
import pandas as pd #load the pandas library 

#Loading Raw Data 
file_path = '/Users/adelfio/Library/CloudStorage/OneDrive-Personal/CV & Work/Projects/AO Project/2024.xlsx' #Define the file path to excel document
xls = pd.ExcelFile(file_path) # 'pd.ExcelFile' allows access to multiple sheets inside the file, doesnt load all the data- just the structure
                              #  df- stores the excel file object

#Cleaning and Preprocessing the Data 
def clean_sheet(sheet_name):
    df = xls.parse(sheet_name) #Reads and converts the sheets into  a structured df

    df.columns = df.iloc[0] # Sets first row as column name (headers) / df.column changes the column names / 'df.iloc[0]' converts row 0 into a 1D Series (column)
    df = df[1:].reset_index(drop=True) # [1:] drops the first row / .reset_index(drop=True) Resets the new index back to 0
    df = df.dropna(axis=1, how='all') #Drops all empty series from the df / 'axis = 1' refers to columns / 'how = all'- only drop column if its missing all values 
    #df = df.iloc[1:, :27] #Pandas method 'integer location'. Excludes the first row and uses columns 0 - 27
    
    df = df.rename(columns={ #renaming important columns 
        df.columns[0]: "Category",
        df.columns[1]: "Product",
        df.columns[2]: "Cost Price"
})
    
    df["Category"] = df["Category"].ffill()  #infilling missing category values #'ffill' fills in missing value with last present value 
    df["Cost Price"] = pd.to_numeric(df["Cost Price"], errors='coerce') #Convert Cost Price to Numeric Values #If value isnt readable/present, 'coerce' converts missing/wrongvalues to NaN instead of causing an error
    
    kitchen_index = df.columns.get_loc("Kitchen") #Retrieves the number of the  'Kitchen' column
    #bar_columns = df.columns[3:kitchen_index] # only using columns from 3 up to, not including 'Kitchen'  
    columns_list = df.columns.tolist()
    kitchen_index = columns_list.index("Kitchen")
    bar_columns = columns_list[3:kitchen_index]

    if "Kitchen" not in df.columns:
        raise ValueError(f"'Kitchen' column not found in sheet: {sheet_name}")

    id_vars = df[["Category", "Product", "Cost Price"]] #Copys these columns preserving them as is 
    df_bars = pd.concat([id_vars, df[bar_columns]], axis=1) #pd.concat is for combining multiple df's or Series- vertically (axis = 0) or horizontally (axis = 1)

    melt_columns = df_bars.columns[3:]    # Create a column list for melting  #melt converts the rows the columns 

    # Melt wide → long
    df_long = df_bars.melt( #new df created
        id_vars=["Category", "Product", "Cost Price"], # columns to keep the same 
        value_vars=melt_columns, # columns to melt
        var_name="Bar", # Column thats turned to rows
        value_name="Quantity" #values assigned to new column
    )
    print("Melting bar columns:", bar_columns)


    df_long["Quantity"] = pd.to_numeric(df_long["Quantity"], errors="coerce")  # Convert quantity to numeric
    df_long["Month"] = sheet_name   # Add the sheet name as month

    return df_long

all_sheets =[] #New list with variable names 'all_sheets' used to store multiple df's (the multiple months sheets)
for sheet in xls.sheet_names: #Loops through all the sheets 'for'- loop, 'sheet'- loop variable, 'in xls.sheet_names'- in the excel sheet 
    print(f"Processing  sheet: {sheet}") #using 'f-string' 
    cleaned = clean_sheet(sheet)
    all_sheets.append(cleaned) 

full_df = pd.concat(all_sheets, ignore_index = True)



ordered_columns = ["Month", "Category", "Product", "Bar", "Cost Price", "Quantity"]
full_df = full_df[ordered_columns]

full_df = full_df.dropna(subset=["Quantity", "Cost Price"])
full_df = full_df[full_df["Quantity"] > 0]
full_df.to_excel("AO_Combined_Cleaned.xlsx", index=False)
print("✅ Excel file saved as AO_Combined_Cleaned.xlsx")


#ordered_columns = ["Month", "Category", "Product", "Bar", "Cost Price", "Quantity", "Cost per Bar"]
#full_df["Cost per Bar"] = full_df["Cost Price"] * full_df["Quantity"] 
#full_df["Cost per Bar"] = full_df["Cost per Bar"].round(2)
# %%
