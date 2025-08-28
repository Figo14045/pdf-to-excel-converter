import streamlit as st
import pandas as pd
import pdfplumber
import io
import re
from datetime import datetime

# Page configuration
st.set_page_config(
    page_title="PDF to Excel Converter",
    page_icon="📊",
    layout="centered"
)

class SimplePDFConverter:
    def __init__(self):
        self.tables = []
        self.metadata = {}
    
    def extract_tables_from_pdf(self, pdf_file):
        """Extract all tables from PDF file"""
        try:
            with pdfplumber.open(pdf_file) as pdf:
                full_text = ""
                all_tables = []
                
                # Extract text and tables from all pages
                for page_num, page in enumerate(pdf.pages):
                    # Get text
                    page_text = page.extract_text()
                    if page_text:
                        full_text += page_text + "\n"
                    
                    # Extract tables
                    tables = page.extract_tables()
                    if tables:
                        for table_idx, table in enumerate(tables):
                            if table and len(table) > 1:  # Ensure table has content
                                all_tables.append({
                                    'page': page_num + 1,
                                    'table_idx': table_idx + 1,
                                    'data': table
                                })
                
                # Clean and process tables
                self.tables = []
                for table_info in all_tables:
                    cleaned_data = self.clean_table_data(table_info['data'])
                    if cleaned_data:
                        self.tables.append({
                            'name': f"Page_{table_info['page']}_Table_{table_info['table_idx']}",
                            'data': cleaned_data,
                            'rows': len(cleaned_data) - 1,  # Exclude header
                            'columns': len(cleaned_data[0]) if cleaned_data else 0
                        })
                
                # Extract basic metadata
                self.metadata = self.extract_basic_info(full_text, pdf_file.name)
                
                return True, f"Successfully extracted {len(self.tables)} tables"
                
        except Exception as e:
            return False, f"Error processing PDF: {str(e)}"
    
    def clean_table_data(self, raw_table):
        """Clean and standardize table data"""
        if not raw_table or len(raw_table) < 2:
            return None
        
        cleaned_table = []
        for row in raw_table:
            if row and any(cell and str(cell).strip() for cell in row):
                # Clean each cell
                cleaned_row = []
                for cell in row:
                    if cell is None:
                        cleaned_row.append('')
                    else:
                        # Clean whitespace and formatting
                        cell_str = str(cell).strip()
                        cell_str = re.sub(r'\s+', ' ', cell_str)
                        cleaned_row.append(cell_str)
                cleaned_table.append(cleaned_row)
        
        return cleaned_table if len(cleaned_table) > 1 else None
    
    def extract_basic_info(self, text, filename):
        """Extract basic document information"""
        info = {
            'filename': filename,
            'processed_at': datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        }
        
        # Try to detect if it's a Shopee document
        if any(keyword in text.lower() for keyword in ['shopee', 'income statement', 'payout']):
            info['document_type'] = 'Shopee Income Statement'
            
            # Extract company name
            company_match = re.search(r'Name in Bank Account\s*:\s*([^\n]+)', text)
            if company_match:
                info['company'] = company_match.group(1).strip()
            
            # Extract period
            period_match = re.search(r'Statement for\s+(\d{4}-\d{2}-\d{2})\s+to\s+(\d{4}-\d{2}-\d{2})', text)
            if period_match:
                info['period'] = f"{period_match.group(1)} to {period_match.group(2)}"
        else:
            info['document_type'] = 'General PDF Document'
        
        return info
    
    def create_excel_file(self):
        """Create Excel file with all extracted tables"""
        if not self.tables:
            return None
        
        output = io.BytesIO()
        
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            workbook = writer.book
            
            # Create formatting
            header_format = workbook.add_format({
                'bold': True,
                'bg_color': '#4CAF50',
                'font_color': 'white',
                'border': 1
            })
            
            cell_format = workbook.add_format({
                'border': 1,
                'align': 'left'
            })
            
            # Create document info sheet
            if self.metadata:
                info_data = [[k.replace('_', ' ').title(), v] for k, v in self.metadata.items()]
                df_info = pd.DataFrame(info_data, columns=['Property', 'Value'])
                df_info.to_excel(writer, sheet_name='Document_Info', index=False)
                
                worksheet = writer.sheets['Document_Info']
                worksheet.set_column('A:A', 20)
                worksheet.set_column('B:B', 30)
                worksheet.set_row(0, None, header_format)
            
            # Create sheet for each table
            for table in self.tables:
                try:
                    # Convert to DataFrame
                    df = pd.DataFrame(table['data'][1:], columns=table['data'][0])
                    
                    # Clean sheet name (Excel limitations)
                    sheet_name = table['name'][:31].replace('/', '_').replace('\\', '_')
                    
                    # Write to Excel
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
                    
                    # Format the sheet
                    worksheet = writer.sheets[sheet_name]
                    
                    # Auto-adjust column widths
                    for i, col in enumerate(df.columns):
                        max_len = max(
                            df[col].astype(str).map(len).max() if len(df) > 0 else 0,
                            len(str(col))
                        ) + 2
                        worksheet.set_column(i, i, min(max_len, 50))
                    
                    # Apply formatting
                    worksheet.set_row(0, None, header_format)
                    
                except Exception as e:
                    st.warning(f"Could not process table {table['name']}: {str(e)}")
                    continue
        
        output.seek(0)
        return output

# Streamlit App
def main():
    # Header
    st.title("📊 PDF to Excel Converter")
    st.markdown("Convert your PDF tables to Excel format - **Free & Simple**")
    
    # Instructions
    with st.expander("📋 How to use"):
        st.markdown("""
        1. **Upload your PDF file** using the file uploader below
        2. **Wait for processing** - tables will be automatically detected
        3. **Download the Excel file** with all extracted tables
        4. **Open in Excel/Google Sheets** and use your data!
        
        **Supported:** Shopee statements, invoices, reports, and any PDF with tables.
        """)
    
    # File uploader
    uploaded_file = st.file_uploader(
        "Choose a PDF file",
        type=['pdf'],
        help="Upload a PDF file containing tables to convert to Excel"
    )
    
    if uploaded_file is not None:
        # Display file info
        st.success(f"📁 File uploaded: **{uploaded_file.name}** ({uploaded_file.size / 1024:.1f} KB)")
        
        # Process button
        if st.button("🔄 Convert to Excel", type="primary"):
            with st.spinner("Processing PDF... This may take a moment"):
                # Create converter and process file
                converter = SimplePDFConverter()
                success, message = converter.extract_tables_from_pdf(uploaded_file)
                
                if success:
                    st.success(message)
                    
                    # Display results
                    if converter.metadata:
                        st.subheader("📄 Document Information")
                        col1, col2 = st.columns(2)
                        
                        with col1:
                            st.write(f"**Type:** {converter.metadata.get('document_type', 'N/A')}")
                            st.write(f"**Filename:** {converter.metadata.get('filename', 'N/A')}")
                        
                        with col2:
                            if 'company' in converter.metadata:
                                st.write(f"**Company:** {converter.metadata['company']}")
                            if 'period' in converter.metadata:
                                st.write(f"**Period:** {converter.metadata['period']}")
                    
                    # Show tables summary
                    st.subheader("📊 Extracted Tables")
                    for i, table in enumerate(converter.tables, 1):
                        st.write(f"**Table {i}:** {table['rows']} rows × {table['columns']} columns")
                    
                    # Show table previews
                    for i, table in enumerate(converter.tables):
                        with st.expander(f"Preview: {table['name']}"):
                            try:
                                df_preview = pd.DataFrame(
                                    table['data'][1:6],  # Show first 5 rows
                                    columns=table['data'][0]
                                )
                                st.dataframe(df_preview, use_container_width=True)
                                
                                if table['rows'] > 5:
                                    st.write(f"*... and {table['rows'] - 5} more rows*")
                            except Exception as e:
                                st.write("Could not preview this table")
                    
                    # Create and offer download
                    excel_file = converter.create_excel_file()
                    
                    if excel_file:
                        # Generate filename
                        original_name = uploaded_file.name.replace('.pdf', '')
                        excel_filename = f"{original_name}_converted_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
                        
                        # Download button
                        st.subheader("📥 Download Your Excel File")
                        st.download_button(
                            label="📥 Download Excel File",
                            data=excel_file,
                            file_name=excel_filename,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            help="Click to download the converted Excel file"
                        )
                        
                        st.success("✅ Conversion completed! Click the download button above.")
                    else:
                        st.error("Could not create Excel file. Please try again.")
                
                else:
                    st.error(message)
    
    # Footer
    st.markdown("---")
    st.markdown(
        "<div style='text-align: center; color: gray;'>"
        "Made with ❤️ using Streamlit | Free PDF to Excel Conversion"
        "</div>", 
        unsafe_allow_html=True
    )

if __name__ == "__main__":
    main()