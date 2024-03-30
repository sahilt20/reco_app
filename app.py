import streamlit as st
import pandas as pd
import io
from openpyxl import load_workbook

def perform_reconciliation(bank_statement_df, ledger_books_df):
    # Merge bank statement and ledger books dataframes on common columns
    ledger_books_df.fillna(0,inplace=True)
    bank_statement_df['Date'] = pd.to_datetime(bank_statement_df['Date'])
    ledger_books_df['Date'] = pd.to_datetime(ledger_books_df['Date'])
    merged_df = pd.merge(bank_statement_df, ledger_books_df,on=['Date','Amount withdrawn','Amount deposited'],how='outer', indicator=True)
    print(merged_df)
    # Identify unmatched transactions
    unmatched_bank= merged_df[merged_df['_merge'] == 'left_only']
    unmatched_bank_Amount_withdrawn = unmatched_bank[unmatched_bank['Amount withdrawn'] > 0]
    unmatched_bank_Amount_deposited = unmatched_bank[unmatched_bank['Amount deposited'] > 0]
    
    unmatched_ledger = merged_df[merged_df['_merge'] == 'right_only']
    unmatched_ledger_Amount_withdrawn = unmatched_ledger[unmatched_ledger['Amount withdrawn'] > 0]
    unmatched_ledger_Amount_deposited = unmatched_ledger[unmatched_ledger['Amount deposited'] > 0]
    
    # Calculate total debits and credits
    total_bank_debit = bank_statement_df['Amount withdrawn'].sum()
    total_bank_credit = bank_statement_df['Amount deposited'].sum()
    total_ledger_debit = ledger_books_df['Amount withdrawn'].sum()
    total_ledger_credit = ledger_books_df['Amount deposited'].sum()
    
    # Reconciliation summary
    reconciliation_summary = {
        'Total Bank Debit': total_bank_debit,
        'Total Bank Credit': total_bank_credit,
        'Total Ledger Debit': total_ledger_debit,
        'Total Ledger Credit': total_ledger_credit,
        'Unmatched Bank Transactions': len(unmatched_bank),
        'Unmatched Ledger Transactions': len(unmatched_ledger)
    }
    
    # Annexures
    annexure1_df = unmatched_bank_Amount_withdrawn
    annexure2_df = unmatched_bank_Amount_deposited
    annexure3_df = unmatched_ledger_Amount_withdrawn
    annexure4_df = unmatched_ledger_Amount_deposited
    
    return reconciliation_summary, annexure1_df, annexure2_df, annexure3_df, annexure4_df

def main():
    st.title('Bank Reconciliation App')
    
    uploaded_file = st.file_uploader("Choose an XLSX file", type="xlsx")
    
    if uploaded_file is not None:
        # Load Excel file into a dictionary of DataFrames
        excel_data = pd.read_excel(uploaded_file, sheet_name=None, skiprows=3)
        
        # Extract specific sheets for reconciliation
        bank_statement_df = excel_data.get('BankStatement', pd.DataFrame())
        ledger_books_df = excel_data.get('BankLedger_Books', pd.DataFrame())
        ledger_books_df.drop(ledger_books_df.index[0], inplace=True)
        ledger_books_df = ledger_books_df.drop(ledger_books_df.index[-1])
        
        # Perform reconciliation
        reconciliation_summary, annexure1_df, annexure2_df, annexure3_df, annexure4_df = perform_reconciliation(bank_statement_df, ledger_books_df)
        
        # Display reconciliation summary and annexures
        st.header('Reconciliation Summary')
        st.write(reconciliation_summary)
        
        st.header(f'Annexure 1 : Amount credited in bank, but not in ledger (({len(annexure1_df)}))')
        st.write(annexure1_df)
        
        st.header(f'Annexure 2 : Amount debited in ledger, but not in bank (({len(annexure2_df)}))')
        st.write(annexure2_df)
        
        st.header(f'Annexure 3 : Amount debited in bank, but not in ledger (({len(annexure3_df)}))')
        st.write(annexure3_df)
        
        st.header(f'Annexure 4 : Amount credited in ledger, but not in bank (({len(annexure4_df)}))')
        st.write(annexure4_df)
        
        # Prepare the output Excel file
        output_excel = io.BytesIO()
        
        # Use XlsxWriter to write the reconciliation output
        with pd.ExcelWriter(output_excel, engine='xlsxwriter') as writer:
            # Write the reconciliation summary and annexures
            pd.DataFrame.from_dict(reconciliation_summary, orient='index', columns=['Value']).to_excel(writer, sheet_name='Reconciliation_Summary')
            annexure1_df.to_excel(writer, sheet_name='Annexure1', index=False)
            annexure2_df.to_excel(writer, sheet_name='Annexure2', index=False)
            annexure3_df.to_excel(writer, sheet_name='Annexure3', index=False)
            annexure4_df.to_excel(writer, sheet_name='Annexure4', index=False)
            
            # Write all input sheets to the output file
            for sheet_name, df in excel_data.items():
                df.to_excel(writer, sheet_name=sheet_name, index=False)
        
        # Reset the pointer of the BytesIO object to the beginning
        output_excel.seek(0)
        
        # Provide a download button for the output Excel file
        st.download_button(
            label="Download Reconciliation Excel",
            data=output_excel,
            file_name="reconciliation_output.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

if __name__ == '__main__':
    main()
