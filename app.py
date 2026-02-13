import streamlit as st

st.set_page_config(
    page_title="Avaya Tools Portal",
    page_icon="📞",
)

# Avaya Logo
st.image("https://upload.wikimedia.org/wikipedia/commons/thumb/e/e0/Avaya_logo.svg/512px-Avaya_logo.svg.png", width=200)

st.write("# Welcome to the Avaya Tools Portal! 👋")

# Internal Use Warning Banner
st.warning("🔒 **Internal Use Only:** This tool is for authorized team members to process Avaya configuration files. Please ensure you do not upload external or unauthorized data.")

# Instructions Section
st.markdown("### 📖 How to use this portal")
st.markdown(
    """    
    **Your Step-by-Step Workflow:**
    
    **1️⃣ Extract Buttons (Format for Editing)**
    * Select **1 Extract Buttons** from the sidebar on the left.
    * Upload your raw `Endpoints` file.
    * Download the resulting Excel file, where all the buttons are cleanly separated into editable columns.
    
    **2️⃣ Make Your Edits**
    * Open the downloaded file in Excel.
    * Make your necessary button changes or updates.
    * Save the file to your computer.
    
    **3️⃣ Reconstruct ACCEC (Format for Import)**
    * Select **2 Reconstruct ACCEC** from the sidebar on the left.
    * Upload your modified file.
    * Download the final file, which is now properly formatted and ready for ACCEC import!
    
    ---
    *💡 **Troubleshooting Tip:** If you run into an error, double-check that your uploaded Excel files have the correct sheet names (`Endpoints` for step 1, and `Avaya Buttons` for step 3).*
    """
)
