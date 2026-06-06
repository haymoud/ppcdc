import pandas as pd
import numpy as np
from flask import Blueprint, render_template, request, send_file, flash, redirect
from io import BytesIO
import warnings

warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

# Création du Blueprint 'balance'
balance_bp = Blueprint('balance', __name__, template_folder='templates')

ALLOWED_EXTENSIONS = {'xls', 'xlsx'}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@balance_bp.route('/')
def index():
    return render_template('balance_index.html')

@balance_bp.route('/process', methods=['POST'])
def process_journal():
    if 'file' not in request.files:
        flash("Aucun fichier n'a été envoyé.")
        return redirect('/balance') # Redirection vers la route du blueprint
    
    file = request.files['file']
    
    if file.filename == '':
        flash('Aucun fichier sélectionné.')
        return redirect('/balance')
        
    if file and allowed_file(file.filename):
        try:
            # Pandas utilisera automatiquement xlrd pour les .xls et openpyxl pour les .xlsx
            df_journal = pd.read_excel(file, sheet_name=0, header=6)
            
            accounts = df_journal.iloc[:, 3].fillna(df_journal.iloc[:, 4])
            debit_amounts = df_journal.iloc[:, 6].fillna(0)
            credit_amounts = df_journal.iloc[:, 7].fillna(0)
            
            valid_mask = accounts.notna()
            parsed_df = pd.DataFrame({
                'N° Compte': accounts[valid_mask],
                'Débit': debit_amounts[valid_mask],
                'Crédit': credit_amounts[valid_mask]
            })
            
            parsed_df['Débit'] = pd.to_numeric(parsed_df['Débit'], errors='coerce').fillna(0)
            parsed_df['Crédit'] = pd.to_numeric(parsed_df['Crédit'], errors='coerce').fillna(0)
            
            def format_account(x):
                try:
                    return str(int(float(x)))
                except:
                    return str(x).strip()
                    
            parsed_df['N° Compte'] = parsed_df['N° Compte'].apply(format_account)
            
            balance = parsed_df.groupby('N° Compte', as_index=False)[['Débit', 'Crédit']].sum()
            balance['Solde Net'] = balance['Débit'] - balance['Crédit']
            balance['Solde Débiteur'] = balance['Solde Net'].apply(lambda x: x if x > 0 else 0)
            balance['Solde Créditeur'] = balance['Solde Net'].apply(lambda x: abs(x) if x < 0 else 0)
            
            for col in ['Débit', 'Crédit', 'Solde Débiteur', 'Solde Créditeur']:
                balance[col] = balance[col].replace(0, np.nan)
                
            balance.drop(columns=['Solde Net'], inplace=True)
            
            totals = pd.DataFrame({
                'N° Compte': ['TOTAL'],
                'Débit': [balance['Débit'].sum()],
                'Crédit': [balance['Crédit'].sum()],
                'Solde Débiteur': [balance['Solde Débiteur'].sum()],
                'Solde Créditeur': [balance['Solde Créditeur'].sum()]
            })
            balance = pd.concat([balance, totals], ignore_index=True)

            output = BytesIO()
            writer = pd.ExcelWriter(output, engine='xlsxwriter')
            balance.to_excel(writer, sheet_name='Balance Générée', index=False)
            
            workbook  = writer.book
            worksheet = writer.sheets['Balance Générée']
            
            header_format = workbook.add_format({
                'bold': True, 'text_wrap': True, 'valign': 'top', 
                'fg_color': '#2C3E50', 'font_color': 'white', 'border': 1
            })
            money_format = workbook.add_format({'num_format': '#,##0.00', 'border': 1})
            account_format = workbook.add_format({'align': 'left', 'border': 1, 'bold': True})
            total_format = workbook.add_format({'bold': True, 'num_format': '#,##0.00', 'bg_color': '#ECF0F1', 'border': 1})
            
            for col_num, value in enumerate(balance.columns.values):
                worksheet.write(0, col_num, value, header_format)
                
            worksheet.set_column('A:A', 18, account_format)
            worksheet.set_column('B:E', 20, money_format)
            
            total_row_idx = len(balance)
            worksheet.write(total_row_idx, 0, 'TOTAL', workbook.add_format({'bold': True, 'border': 1, 'bg_color': '#ECF0F1'}))
            worksheet.write(total_row_idx, 1, balance.iloc[-1]['Débit'], total_format)
            worksheet.write(total_row_idx, 2, balance.iloc[-1]['Crédit'], total_format)
            worksheet.write(total_row_idx, 3, balance.iloc[-1]['Solde Débiteur'], total_format)
            worksheet.write(total_row_idx, 4, balance.iloc[-1]['Solde Créditeur'], total_format)
            
            writer.close()
            output.seek(0)
            
            return send_file(
                output,
                as_attachment=True,
                download_name="Balance_Comptable_Generee.xlsx",
                mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
        except Exception as e:
            flash(f"Erreur de lecture : Vérifiez que le fichier est valide. Détail: {str(e)}")
            return redirect('/balance')
    else:
        flash('Format de fichier non autorisé. Veuillez utiliser .xls ou .xlsx')
        return redirect('/balance')