import pandas as pd
import matplotlib.pyplot as plt
import os
import xlsxwriter

def load_data(file_path):
    file_extension = os.path.splitext(file_path)[1].lower()
    if file_extension == '.ods':
        data = pd.read_excel(file_path, engine='odf')
    elif file_extension == '.xlsx' or file_extension == '.xls':
        data = pd.read_excel(file_path)
    else:
        raise ValueError("Unsupported file format. Please provide an ODS or Excel file.")
    return data

def clean_data(data):
    # Skip rows until the header is found
    data = data.iloc[2:]
    data.columns = data.iloc[0]
    data = data[1:]
    
    # Extract relevant columns and remove rows with NaN values in 'Profil', 'Qté', or 'Long.'
    data_cleaned = data[['Profil', 'Qté', 'Long.']].dropna(subset=['Profil', 'Qté', 'Long.'])
    data_cleaned = data_cleaned[~data_cleaned['Profil'].str.contains('Total')]
    data_cleaned['Qté'] = data_cleaned['Qté'].astype(int)
    data_cleaned['Long.'] = data_cleaned['Long.'].astype(int)
    return data_cleaned

def optimize_cutting(data, stock_length):
    results = []
    
    for profile, group in data.groupby('Profil'):
        pieces = []
        
        for _, row in group.iterrows():
            pieces.extend([row['Long.']] * row['Qté'])
        
        # Sort pieces in descending order
        pieces.sort(reverse=True)
        
        stock_used = []
        current_stock = stock_length
        current_pieces = []

        for piece in pieces:
            if piece <= current_stock:
                current_pieces.append(piece)
                current_stock -= piece
            else:
                stock_used.append((current_pieces, stock_length - current_stock))
                current_pieces = [piece]
                current_stock = stock_length - piece

        if current_pieces:
            stock_used.append((current_pieces, stock_length - current_stock))
        
        results.append((profile, stock_length, stock_used))
    
    return results

def draw_cutting_plan(results, save_path):
    fig, axs = plt.subplots(len(results), figsize=(10, len(results) * 2))
    
    if len(results) == 1:
        axs = [axs]
    
    for i, (profile, stock_length, stock_used) in enumerate(results):
        for j, (pieces, total_used) in enumerate(stock_used):
            axs[i].barh(y=j, width=stock_length, color='grey', edgecolor='black')
            start = 0
            for piece in pieces:
                axs[i].barh(y=j, width=piece, left=start, color='blue', edgecolor='black')
                start += piece
            axs[i].set_xlim(0, stock_length)
            axs[i].set_title(f'Profile {profile}: Piece {j+1}: Total Length Used = {total_used} mm, Remaining = {stock_length - total_used} mm')
            axs[i].axis('off')

    plt.tight_layout()
    plt.savefig(save_path)
    plt.close()

def export_to_excel(results, output_path, image_path):
    data_to_export = []
    for profile, stock_length, stock_used in results:
        for j, (pieces, total_used) in enumerate(stock_used):
            data_to_export.append([profile, stock_length, j+1, pieces, total_used, stock_length - total_used])
    
    df = pd.DataFrame(data_to_export, columns=['Profile', 'Stock Length', 'Cut Index', 'Pieces Cut', 'Total Length Used', 'Remaining Length'])
    
    writer = pd.ExcelWriter(output_path, engine='xlsxwriter')
    df.to_excel(writer, sheet_name='Cutting Plan', index=False)
    
    workbook  = writer.book
    worksheet = writer.sheets['Cutting Plan']
    
    worksheet.insert_image('H2', image_path)
    
    writer._save()

def main(file_path, output_path, image_path, stock_length):
    data = load_data(file_path)
    data_cleaned = clean_data(data)
    results = optimize_cutting(data_cleaned, stock_length)
    draw_cutting_plan(results, image_path)
    export_to_excel(results, output_path, image_path)
    
    for profile, stock_length, stock_used in results:
        total_stock_used = len(stock_used)
        total_length_saved = stock_length * total_stock_used - sum(data_cleaned[data_cleaned['Profil'] == profile]['Long.'] * data_cleaned[data_cleaned['Profil'] == profile]['Qté'])
        print(f"Profile {profile} requires {total_stock_used} pieces of {stock_length} mm stock length. Total length saved: {total_length_saved} mm.")

# Example usage
file_path = 'A_liste appro.xlsx'
output_path = 'optimized_cutting_plan.xlsx'
image_path = 'cutting_plan.png'
stock_length = 12000
main(file_path, output_path, image_path, stock_length)
