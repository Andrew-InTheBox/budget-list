
# Protocol for Setting Up Dependent Dropdowns in Excel (Categories and Subcategories)

## Step 1: Create a Reference Sheet
1. Organize your **categories** and their corresponding **subcategories** in two columns:
   - **Column A**: Categories (e.g., Travel, Utilities).
   - **Column B**: Subcategories (e.g., Hotel, Flight under Travel).

## Step 2: Create Named Ranges for Subcategories
1. Select the subcategories for each category in the **Reference** sheet.
2. Go to **Formulas** > **Define Name**.
3. Name the range `<Category>_Subcategories` (e.g., `Travel_Subcategories` for Travel).
4. Repeat for all categories.

## Step 3: Set Up Category Dropdown in Column A
1. In your **Main Sheet**, select the range where categories will be chosen (e.g., `A2:A100`).
2. Go to **Data** > **Data Validation**.
3. Set the **Allow** option to `List` and the **Source** to:
   ```excel
   =UNIQUE(Reference!A:A)
   ```
4. Click **OK**.

## Step 4: Set Up Dependent Dropdown in Column B
1. In your **Main Sheet**, select the range where subcategories will be chosen (e.g., `B2:B100`).
2. Go to **Data** > **Data Validation**.
3. Set the **Allow** option to `List` and the **Source** to:
   ```excel
   =INDIRECT(A2 & "_Subcategories")
   ```
4. Click **OK**.

## Step 5: Test the Dropdowns
1. In Column A, select a category from the dropdown.
2. Verify that Column B shows the corresponding subcategories.

---

## Key Notes:
- **Named Ranges** must match the category names exactly, appended with `_Subcategories` (e.g., `Travel_Subcategories`).
- **Blank Cells**: Ensure your Reference sheet has no blank rows within the data range.
- **Dynamic Updates**: To update categories or subcategories, modify the Reference sheet and adjust Named Ranges as needed.

This protocol ensures a clean setup for dependent dropdowns. Let me know if you need further assistance!
