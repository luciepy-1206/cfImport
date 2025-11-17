        ### Instructions:
        
        1. **Select Mode**:
           - **Create New File**: Creates a new Excel file with empty sheets and applies CF rules
           - **Update Existing File**: Opens an existing Excel file and applies CF rules to it
        
        2. **Upload Files**:
           - Upload your CF Rules file (required)
           - Upload target file (only if updating existing file)
        
        3. **Configure Range**:
           - **First Row**: The starting row number (e.g., 13)
           - **Last Row**: The ending row number (e.g., 15)
           - This creates ranges like `AE13:AJ15` instead of `AE1:AJ1000`
        
        4. **Configure @ROW@ Replacement**:
           - **Leave empty**: Uses the First Row value to replace @ROW@
           - **Enter a number**: Uses that specific number to replace @ROW@
        
        5. **Generate**: Click the "Generate Excel File with CF Rules" button
        
        6. **Download**: Click the download button to get your formatted Excel file
        
        ### Examples:
        
        **Example 1 - Create New:**
        - Mode: Create New File
        - First Row: 13, Last Row: 15
        - Row Number for @ROW@: (empty)
        - Result: New file with rules applied to rows 13-15, @ROW@ = 13
        
        **Example 2 - Update Existing:**
        - Mode: Update Existing File
        - Upload: Your existing workbook
        - First Row: 2, Last Row: 100
        - Row Number for @ROW@: 2
        - Result: Your file with CF rules applied, preserving existing data
        
        ### Requirements:
        
        Your CF Rules file must have a sheet named **"CF Rules"** with these columns:
        - Start Column, End Column, Formula, Stop if True
        - BG Color, BG RGB, Font Color, Font RGB
        - Number Format, Worksheet Name
        """)

