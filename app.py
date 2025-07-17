import streamlit as st
import os
import pandas as pd
from agno_agent import create_sync_agent
from agno_ppt_excel_agent import ExtractExcelData, ExtractPPTText
import logging
import tempfile
import shutil

# Set up logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Configuration
UPLOAD_DIR = "uploads"
os.makedirs(UPLOAD_DIR, exist_ok=True)

# Page config
st.set_page_config(
    page_title="Excel to PowerPoint Sync AI",
    page_icon="📊",
    layout="wide"
)

st.title("📊 Excel to PowerPoint Sync AI")
st.markdown("Upload your Excel/CSV and PowerPoint files to automatically sync data changes.")

# Create columns for file uploads
col1, col2 = st.columns(2)

with col1:
    st.subheader("📄 PowerPoint File")
    ppt_file = st.file_uploader("Upload PowerPoint file", type=["pptx"])
    if ppt_file:
        st.success(f"✅ Loaded: {ppt_file.name}")

with col2:
    st.subheader("📊 Excel/CSV File")
    excel_file = st.file_uploader("Upload Excel/CSV file", type=["xlsx", "xls", "csv"])
    if excel_file:
        st.success(f"✅ Loaded: {excel_file.name}")

# Process files
if st.button("🔄 Sync PowerPoint with Excel Data", type="primary"):
    if ppt_file and excel_file:
        try:
            # Create progress bar
            progress_bar = st.progress(0)
            status_text = st.empty()
            
            # Save uploaded files
            status_text.text("💾 Saving uploaded files...")
            progress_bar.progress(10)
            
            ppt_path = os.path.join(UPLOAD_DIR, ppt_file.name)
            excel_path = os.path.join(UPLOAD_DIR, excel_file.name)
            output_path = os.path.join(UPLOAD_DIR, f"updated_{ppt_file.name}")
            
            # Save files
            with open(ppt_path, "wb") as f:
                f.write(ppt_file.getvalue())
            with open(excel_path, "wb") as f:
                f.write(excel_file.getvalue())
            
            # Convert CSV to Excel if needed
            if excel_file.name.lower().endswith('.csv'):
                status_text.text("🔄 Converting CSV to Excel...")
                progress_bar.progress(20)
                
                df = pd.read_csv(excel_path)
                excel_path = excel_path.replace('.csv', '.xlsx')
                df.to_excel(excel_path, index=False)
            
            # Extract data
            status_text.text("📊 Extracting Excel data...")
            progress_bar.progress(30)
            
            excel_extractor = ExtractExcelData()
            excel_data = excel_extractor.run(excel_path)
            
            if "error" in excel_data:
                st.error(f"❌ Error extracting Excel data: {excel_data['error']}")
                st.stop()
            
            # Extract PPT text
            status_text.text("📄 Analyzing PowerPoint content...")
            progress_bar.progress(50)
            
            ppt_extractor = ExtractPPTText()
            ppt_data = ppt_extractor.run(ppt_path)
            
            if "error" in ppt_data:
                st.error(f"❌ Error extracting PowerPoint text: {ppt_data['error']}")
                st.stop()
            
            # Create and run agent
            status_text.text("🤖 Running AI synchronization...")
            progress_bar.progress(70)
            
            agent = create_sync_agent()
            
            # Prepare the prompt
            prompt = f"""
Please synchronize the PowerPoint presentation with the Excel data:

PowerPoint file: {ppt_path}
Excel data extracted: {len(excel_data.get('numbers', []))} numbers, {len(excel_data.get('text_values', []))} text values, {len(excel_data.get('key_value_pairs', {}))} key-value pairs
PowerPoint content: {len(ppt_data.get('slides', []))} slides with {len(ppt_data.get('numbers', []))} numbers found

Output path: {output_path}

Please update the PowerPoint with the new Excel data while preserving all formatting and layout.
"""
            
            # Run the agent
            try:
                response = agent.run(prompt)
                
                status_text.text("✅ Processing complete!")
                progress_bar.progress(100)
                
                # Check if output file exists
                if os.path.exists(output_path):
                    st.success("🎉 PowerPoint synchronization completed successfully!")
                    
                    # Display results
                    col1, col2 = st.columns(2)
                    
                    with col1:
                        st.subheader("📊 Excel Data Summary")
                        st.write(f"• Numbers found: {len(excel_data.get('numbers', []))}")
                        st.write(f"• Text values: {len(excel_data.get('text_values', []))}")
                        st.write(f"• Key-value pairs: {len(excel_data.get('key_value_pairs', {}))}")
                        
                        if excel_data.get('key_value_pairs'):
                            st.write("**Key-Value Pairs:**")
                            for key, value in list(excel_data['key_value_pairs'].items())[:5]:
                                st.write(f"• {key}: {value}")
                    
                    with col2:
                        st.subheader("📄 PowerPoint Analysis")
                        st.write(f"• Slides processed: {len(ppt_data.get('slides', []))}")
                        st.write(f"• Numbers found: {len(ppt_data.get('numbers', []))}")
                        st.write(f"• Text elements: {len(ppt_data.get('text_runs', []))}")
                    
                    # Show agent response
                    st.subheader("🤖 AI Agent Response")
                    st.write(response.content)
                    
                    # Download button
                    with open(output_path, "rb") as f:
                        st.download_button(
                            label="📥 Download Updated PowerPoint",
                            data=f.read(),
                            file_name=f"updated_{ppt_file.name}",
                            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                        )
                else:
                    st.error("❌ Output file was not created. Please check the logs.")
                    st.write("**Agent Response:**")
                    st.write(response.content)
                    
            except Exception as e:
                st.error(f"❌ Error running AI agent: {str(e)}")
                logger.error(f"Agent error: {str(e)}")
                
        except Exception as e:
            st.error(f"❌ Error processing files: {str(e)}")
            logger.error(f"Processing error: {str(e)}")
            
    else:
        st.warning("⚠️ Please upload both PowerPoint and Excel/CSV files")

# Sidebar with information
st.sidebar.header("ℹ️ How it works")
st.sidebar.markdown("""
1. **Upload Files**: Select your PowerPoint (.pptx) and Excel/CSV files
2. **AI Analysis**: The AI extracts data from both files
3. **Intelligent Matching**: Numbers and text are matched intelligently
4. **Preserve Formatting**: Original layout and formatting are maintained
5. **Download**: Get your updated PowerPoint file

**Supported Updates:**
- Number updates (with formatting preservation)
- Text replacements
- Key-value pair updates
- Contextual matching
""")

st.sidebar.header("🔧 Features")
st.sidebar.markdown("""
- **Smart Matching**: AI decides what to update
- **Format Preservation**: Keeps $, %, commas, etc.
- **Multi-sheet Support**: Handles multiple Excel sheets
- **Error Handling**: Robust error management
- **Progress Tracking**: Real-time status updates
""")