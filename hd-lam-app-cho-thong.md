**Step by Step Guide on Updating Your SMARTS Data Tracker**

Last updated 3/11/25

Preliminary note: remember to save often and/or have auto save on.

1. **Download the data**

   1. Go to SMARTS  
      https://smarts.waterboards.ca.gov/smarts/SwPublicUserMenu.xhtml

   2. Click on “Public User Menu”

   3. Click on “Download NOI Data By Regional Board”

   4. Select your region from the dropdown menu

   5. Click on both “Industrial Application Specific Data” and “Industrial Ad Hoc Reports \- Parameter Data”

   6. Data will be downloaded to two separate .txt files, each titled “file”

2. **Add the new data to your tracker**

   1. In your existing SMARTS data tracker document, create 2 new sheets with the “+” button on the bottom bar

   2. Get Sheet2 into the proper format for your tracker

      1) Copy (CTRL+C) and paste (CTRL+V) all of the “Industrial Application Specific Data” text file into the first cell in Sheet2 (A1)

      2) Click Row 1 (the header row), go to the “Data” tab in Excel, and click the “Filter” button

      3) Click Column B (the Application ID column), go to the “Home” tab in Excel, and from the “Condition Formatting” button, click “Highlight Cell Rules” then “Duplicate Values” from the drop down menus that appear. Hit “OK” on the textbox that appears.

      4) Go the drop down for Column B and click “filter by color” and then the colored box. This will show only the duplicate rows.

      5) Click on the drop down for Column D (the status column) and uncheck “Active” so that only rows with a status other than “Active” show.

      6) Delete all rows that are showing.

      7) In the “Data” tab in Excel, hit “Clear” to see the rows that are left.

      8) Re-order the columns and delete excess columns to match your tracker

         1) Delete Columns AF (Receiving Water Name) through AL (which should be the right-most column with text)

         2) Delete Columns P (Facility Latitude) through AB (Percent of Site Imperviousness)

         3) Simultaneously delete Columns A, E-H, and L

         4) Move Column B (WDID) to the left of Column A (App ID). Move Column E (Facility Name) to the left of Column D (Operator Name)

         5) After re-ordering, Sheet2 columns should look like the following: A – WDID; B – App ID; C – Status; D – Facility Name; E – Operator Name; F –Address; G – City; H – State; I – Zip; J – Primary SIC; K – Secondary SIC; L – Tertiary SIC

   3. Get Sheet1 into the proper format for your tracker

      1) Copy (CTRL+C) and paste (CTRL+V) all of the “Industrial Ad Hoc Reports \- Parameter Data” text file into the first cell in Sheet1 (A1)

      2) Go to the “Data” tab in Excel, and click the “Remove Duplicates” button and hit “OK” on the box that pops up

      3) In column B (WDID), filter for results showing “4 56” (note the space between 4 and 56), and then delete all rows showing. “4 56” is the county designation for WDID’s for ventura county, and we, therefore, won’t target any of those facilities.

      4) In the “Data” tab in Excel, hit “Clear” to see the rows that are left.

      5) Delete all columns not in your tracker (columns A, J, K, U, X, and Y)

         1) After deleting these columns, Sheet1 columns should look like the following: A – WDID; B – App ID; C – Status; D – Facility Name; E – Operator Name; F –Address; G – City; H – State; I – Zip; J – Primary SIC; K – Secondary SIC; L – Tertiary SIC

   4. Filter Sheet1 for only new sample data

      1) Right click Column B (Application ID), and hit “insert” which will insert a blank column to the left

      2) In the first blank cell (B2), write the following formula and then hit enter: \=VLOOKUP(J2,Data\!O:P,2,FALSE)

      3) If you move your mouse over the bottom right corner of B2, you will see the white plus sign cursor will turn into a thin black plus sign. When that happens, double click and Excel will fill in the formula for the rest of the column.

      4) Then click on the column, hit CTRL+C to copy, and then right click and under “paste options”, hit the icon that has a little “123” in the corner, which will just paste the answers instead of the formulas.

      5) Sort column B from smallest to largest and then from the dropdown for Column B that shows all the numbers, scroll to the bottom and unclick “N/A”. This will leave showing only cells that have numbers in them.

      6) Delete all rows that are showing.

      7) In the “Data” tab in Excel, hit “Clear” to see the rows that are left.

   5. Check if facilities in Sheet1 are active

      1) In Sheet1 cell B2, delete the “N/A” and write the following formula and then hit enter: \=VLOOKUP(C2,Sheet2\!B:D,2,FALSE)

      2) If you move your mouse over the bottom right corner of B2, you will see the white plus sign cursor will turn into a thin black plus sign. When that happens, double click and Excel will fill in the formula for the rest of the column.

      3) Then click on the column, hit CTRL+C to copy, and then right click and under “paste options”, hit the icon that has a little “123” in the corner, which will just paste the answers instead of the formulas.

      4) Sort column B from A to Z and then from the dropdown for Column B that shows the status options, unclick “active”. This will leave showing only cells that are facilities that are not active

      5) Delete all rows that are showing.

      6) In the “Data” tab in Excel, hit “Clear” to see the rows that are left.

   6. Choose the parameters to track in Sheet1

      1) In Sheet1, delete Column B.

      2) Right click column N (the parameter column) and hit insert, which should insert a blank column to the left of column N

      3) From the dropdown for the parameter column (which should now be column O), click the check box next to “(Select All”) which should uncheck all of the parameters. Then go through and check off each parameter that you want to have in your tracker, and then hit “OK”

         1) The parameters LAW has tracked to date are the following (more may be added on a facility-specific basis as needed): Aluminum; Ammonia; Arsenic; Biochemical Oxygen Demand (BOD); Cadmium; Chemical Oxygen Demand (COD); Copper; Cyanide; E. coli; Enterococci MPN; Fecal Coliform; Iron; Lead; Magnesium; Mercury; Nickel; Nitrate; Nitrite; Nitrite plus Nitrate (N+N); Oil & Grease (O\&G); pH; Phosphorus; Selenium; Silver; Total Coliform; Total Suspended Solids (TSS); Zinc

      4) In the first blank cell in column N (the blank column inserted to the left of the parameter column) write the word “keep”, click the cell, and then fill that in to all other cells showing by double clicking the bottom right corner when you see the thin black plus sign.

      5) In the “Data” tab in Excel, hit “Clear” to see the rows that are left.

      6) Sort column N (the one with “keep”) A to Z and in the drop down unclick “keep” so that all that is left are the rows with nothing written.

      7) Delete all rows that are showing.

      8) In the “Data” tab in Excel, hit “Clear” to see the rows that are left.

      9) Delete column N (the one with “keep”)

      10) Right click on what is now column O (“RESULT\_QUALIFIER”) and click “insert” which should insert a blank column to the left of column O

      11) Click on column N (the parameter column) and then in the “Data” tab in Excel, click on “text to columns”.

      12) Make sure that “delimited” is chosen and then click “Next”. 

      13) In the next box, make sure that only “comma” is checked and then hit “finish” (not “next”). This will put everything after the comma (e.g., “dissolved,” “total”, or “total recoverable”) into column O and leave only the parameter names in column N.

   7. Make sure all the samples in Sheet1 are in “mg/L” and not “ug/L”

      1) In Sheet1, sort column R (Units) from A to Z

      2) In the dropdown menu for column R, uncheck everything and then only click ug/L (if there are none, skip the rest of this step)

      3) In the first blank cell to the right of the Reporting Limit column (which should be Column T), write this formula: \=\[first cell in column Q (Result)\]/1000

      4) When the cursor in the bottom right corner of the cell where you wrote that formula turns into the thin black plus sign, click and drag three columns to the right (i.e., there will be four cells in that row with writing in them)

      5) Then, with all four of these cells highlighted, click the thin black plus sign in the bottom right corner of the cell all the way to the right to fill in these four columns for the remaining cells

      6) Without clicking, hit CTRL+C to copy all of these newly filled in cells, and then right click on the first cell showing in column Q (the one written in the formula above) and under “paste options”, hit the icon that has a little “123” in the corner, which will just paste the answers instead of the formulas. 

      7) Then in the first cell showing in column R (units) type in “mg/L” and fill this in for the rest of the rows using the thin black plus sign in the bottom right corner of that cell

      8) Delete columns U-X (the extra columns you created)

      9) In the “Data” tab in Excel, hit “Clear” to go back to all the data.

   8. Add facility information from Sheet2 into Sheet1

      1) In Sheet2, delete Column C so that the application ID column is immediately to the left of the facility name column

         1) After doing this, the columns should look like the following: A – WDID; B – App ID; C – Facility Name; D – Operator Name; E –Address; F – City; G – State; H – Zip; I – Primary SIC; J – Secondary SIC; K – Tertiary SIC

      2) In Sheet1, insert 6 columns to the left of column C (Reporting Year).

      3) In cell C2, write the following formula and hit enter: \=VLOOKUP($B2,Sheet2\!$B:$Z,COLUMN(B:B),FALSE)

      4) When the cursor in the bottom right corner of the cell where you wrote that formula turns into the thin black plus sign, click and drag it to fill in all the blank columns to the right (i.e., there will be 6 cells in that row with writing in them)

      5) Then, with all 6 of these cells highlighted, click the thin black plus sign in the bottom right corner of the cell all the way to the right to fill in these 6 columns for the remaining rows

      6) Without clicking, hit CTRL+C to copy all of these newly filled in cells, and then right click on cell C2 (the one where you originally wrote the above formula) and under “paste options”, hit the icon that has a little “123” in the corner, which will just paste the answers instead of the formulas. 

   9. Add in SIC Codes from Sheet2 into Sheet1

      1) In Sheet2, delete columns C-H, so that the application ID column is immediately to the left of the primary SIC code column

         1) After doing this, the columns should look like the following: A – WDID; B – App ID; C – Primary SIC; D – Secondary SIC; E – Tertiary SIC

      2) In Sheet1, in the header row for the three columns immediately to the right of the last filled in column (likely columns AA-AC), write 1, 2, and 3 respectively 

      3) Click Row 1 (the header row), go to the “Data” tab in Excel, and double click the “Filter” button, which should turn the filters off and back on including now the 3 new columns

      4) In cell AA2, write the following formula and hit enter: \=VLOOKUP($B2,Sheet2\!$B:$Z,COLUMN(B:B),FALSE)

      5) When the cursor in the bottom right corner of the cell where you wrote that formula turns into the thin black plus sign, click and drag it to fill in all the blank columns to the right (i.e., there will be 3 cells in that row with writing in them)

      6) Then, with all 3 of these cells highlighted, click the thin black plus sign in the bottom right corner of the cell all the way to the right to fill in these 3 columns for the remaining rows

      7) Without clicking, hit CTRL+C to copy all of these newly filled in cells, and then right click on cell AA2 (the one where you originally wrote the above formula) and under “paste options”, hit the icon that has a little “123” in the corner, which will just paste the answers instead of the formulas. 

      8) In the dropdown for column AC (what we labeled as “3” for the tertiary SIC code), make sure only the box for “0” is checked

      9) Highlight all the cells in that column, right click and click “clear contents”

      10) In the “Data” tab in Excel, hit “Clear” to go back to all the data.

      11) In the dropdown for column AB (what we labeled as “2” for the secondary SIC code), make sure only the box for “0” is checked

      12) Highlight all the cells in that column, right click and click “clear contents”

      13) In the “Data” tab in Excel, hit “Clear” to go back to all the data.

   10. Combine new data from Sheet1 into existing Data tracker

       1) The new data is now ready to be pasted into the main data tracker but first you need to make sure you don’t have any formulas left over that will break the data.

       2) In the “Home” tab in Excel, click on “Find & Select” and click on “Formulas”. If it says there are no formulas present, you can move on. If it finds formulas, just hit CTRL+A to select all, and then copy (CTRL+C) and paste (CTRL+V) the values (with the little “123” in the corner) to get rid of all formulas).

       3) In the main sheet called “Data”, go to the “Old/New” column and click on the cell in row 2\. Make sure that says “Old” and then with the thin black plus sign cursor, double click to fill that in for the rest of the rows.

       4) Now back in Sheet1, copy (CTRL+C) all rows (except the header row), and paste (CTRL+V) them into the first open row at the bottom of the Data sheet.

       5) Highlight and copy (CTRL+C) the filled in cells in the last row of old data (the row you just pasted under), then go down to the last row of new data and click on the cell in the most right column (the tertiary SIC code column) while holding SHIFT to highlight all the new cells. Right click and under “paste options”, hit the icon that has a little paintbrush and percentage sign, which will only paste the formatting.

       6) Then, in the “Old/New” column, write “new” in the first row of new data and with the thin black plus sign cursor, double click to fill that in for the rest of the rows.

       7) You may need to reformat the Borders on the cells after uploading the new data to make sure the format aligns with older data entries and looks more polished.

          1) To do this, click on Column A, and while holding SHIFT, click on Column AD (this should highlight all filled-in cells in the sheeet). Then at the top of the sheet, click on the “Home” tab and go to the Cells box near the right and click the Format dropdown, then click the “Format Cells” option at the bottom.

          2) In the popup box, click the Border tab and you’ll see a chart on the right-hand side showing four sample cell boxes with the word “Text” in them. On the left of that chart, click on the Top, Middle, and Bottom border icons. Now, the chart should show a solid line at the top, middle, and bottom of the sample cells. Then click “Ok” to confirm these changes.

          3) Finally, un-highlight all of the columns and now click on Column AD to highlight only that column. This should be the right-most column with text in it (the Old/New designation). Once that column is highlighted, go to the “Home” tab and look for the Font box on the left, click the Border icon drop-down menu, and select Right Border.

       8) The data is now ready to review, so delete Sheet1 and Sheet2 and make sure to save.

3. **Analyze the new data**

   1. In the “Data” tab in Excel, highlight the entire sheet and hit “Sort”.

   2. In the box that comes up, you will want a multi-level sort that goes like this:

      1) Sort by “Old/New” from A to Z

      2) Then by “Parameter” from A to Z

      3) Then by “Result” from largest to smallest

      4) Once you input these three directions, click “Ok” to confirm the changes.

   3. In the “Old/New” column dropdown, select only “New” to see just the new results

   4. Before starting to highlight cells, go to the dropdown list for column C (facility name) and write down any facilities that you want to look into regardless of if the new sampling is clean or not (e.g., facilities you are already targeting or litigating against, or are in your compliance program).

   5. Identify exceedances

      1) First you have to figure out what the applicable discharge limits are for facilities in your region so you know what threshold to make highlights of exceedances in the Data sheet.

         1) The IGP has Numeric Action Levels (NALs) that apply to all facilities, on an annual average basis or instantaneous maximum basis.

            1) Annual average NALs: Aluminum – 0.75 mg/L; Ammonia – 2.14 mg/L; Arsenic – 0.15 mg/L; BOD – 30 mg/L; Cadmium – 0.0053 mg/L; COD – 120 mg/L; Copper – 0.0332 mg/L; Cyanide – 0.022 mg/L; Iron – 1.0 mg/L; Lead – 0.262 mg/L; Magnesium – 0.064 mg/L; Mercury – 0.0014 mg/L; Nickel – 1.02 mg/L; N+N – 0.68 mg/L; O\&G – 15 mg/L; Phosphorus – 2.0 mg/L; Selenium – 0.005 mg/L; Silver – 0.0183 mg/L; TSS – 100 mg/L; Zinc – 0.26 mg/L

            2) Instantaneous maximum NALs: O\&G – 25 mg/L; pH – less than 6.0 or greater than 9.0; TSS – 400 mg/L

         2) Then you will need to look up specific TMDL-Related Numeric Action Levels (TNALs) and/or Numeric Effluent Limits (NELs) that apply in your region.

            1) TNALs are usually calculated on either an annual average basis or instantaneous maximum basis.

            2) NELs are usually calculated on an instantaneous maximum basis, with a violation defined as two or more exceedances at the same discharge point in the same reporting year.

      2) Then based on the NALs/NELs/TNALs, etc. go through and highlight the samples that are above the respective limit, and write down the names of any facilities you want to look into further as you go

         1) When highlighting samples above a limit, select all cells in a row, but do not select the entire row. Highlighting the entire row will cause the entire row (even unfilled cells to the right of the last column with text, which should be Column AD) to have highlights, and that ends up looking weird if you filter or re-sort cells.

      3) Be sure to write an explanation for your own reference in the “Explanation” sheet so that you can recall how you did things (e.g., we highlight all facilities based on the LA River NELs for simplicity, even though the NELs do not apply to every facility; we use the lowest of the several ammonia NELs for highlighting, etc.)

         1) LAW’s threshold for highlighting instantaneous TNAL/NEL exceedances is as follows: Ammonia – 4.7 mg/L; Cadmium – 0.0031 mg/L; Copper – 0.06749 mg/L; E. coli – 400/100 mL; Enterococci MPN – 104/100 mL; Fecal Coliform – 400/100 mL; Lead – 0.094 mg/L; Nitrate – 1.0 mg/L; Nitrite – 1.0 mg/L; N+N – 1.0 mg/L; Total Coliform – 10000/100 mL; Zinc – 0.159 mg/L

   6. Once you have highlighted all the new data, in the “Data” tab in Excel, hit “Clear” to go back to all the data.

   7. Looking further into a given facility

      1) Now that you have a list of facilities to look into, this is how you sort Excel to make it easier to look into individual facilities

      2) In the “Data” tab in Excel, hit “Sort”.

      3) In the box that comes up, you will want a multi-level sort that goes like this:

         1) Sort by “WDID” from A to Z

         2) Then by “Reporting Year” from smallest to largest

         3) Then by “Parameter” from A to Z

         4) Then by “Result” from largest to smallest

      4) You can now sort to that specific facility by using their WDID, Application ID, or Facility Name

   8. Check to see which facilities are lying in annual reports about sampling all QSEs

      1) Download annual report data from SMARTS

         1) Go to SMARTS and click on “Public User Menu,” then click on “Download NOI Data By Regional Board”

         2) Select your region from the dropdown menu, then click on “Industrial Annual Reports.” Data will be downloaded to as a .txt file called “file”

         3) Create a new sheet in your SMARTS tracker, which may now be labeled as Sheet3 (if not already labeled as Sheet3, you should rename it to be Sheet3 for purposes of this guide).

         4) Copy (CTRL+C) and paste (CTRL+V) all of the “Industrial Annual Reports” text file into the first cell in Sheet3 (A1)

      2) Get Sheet3 into the proper format for your tracker

         1) To reorder the first two columns, cut (CTRL+X) Column B (WDID) and click on Column A (App ID), then right click and select “Cut Copied Cells”

         2) Delete Columns L (Question 4 Answer) through Column AF (Question TMDL Answer). Then delete Column E (Region) through Column I (Question 2 Explanation)

            1) After doing this, the columns should look like the following: A – WDID; B – App ID; C – Report ID; D – Report Year; E – Question 3 Answer; F – Question 3 Explanation

         3) Highlight the entire tracker in the top left (or just click Row 1), go to the “Data” tab, and click the “Filter” button

      3) Delete Question 3 responses of “No”

         1) Filter Column E (Question 3 Answer), uncheck “Select All,” then check “N” and “Blank” responses

         2) Delete all rows of data (other than Row 1\) that appear

         3) Go to the “Data” tab and click Clear in the “Sort & Filter” box

      4) Compare “Yes” responses to Question 3 with actual sampling data

         1) Right click Column B (Application ID), and hit “insert” which will insert a blank column to the left

         2) In the first blank cell (B2), write the following formula and then hit enter: \=VLOOKUP(J2,Data\!O:P,2,FALSE)

         3) If you move your mouse over the bottom right corner of B2, you will see the white plus sign cursor will turn into a thin black plus sign. When that happens, double click and Excel will fill in the formula for the rest of the column.

         4) Then click on the column, hit CTRL+C to copy, and then right click and under “paste options”, hit the icon that has a little “123” in the corner, which will just paste the answers instead of the formulas.

         5) 

      5) \[fill\]

   9. Check which active, non-NEC facilities have never sampled once over the last 10 years  
      1) \[fill\]

4. **Visualize the data using a Pivot Table (Optional once a specific facility is identified for further investigation or enforcement)**

   1. Once a facility is identified that appears worth looking further into, it is often helpful to use this tool to visualize the data more clearly. Also, once enforcing against the facility, this will be a helpful way to have the data more easily accessible without having to get it from the overall tracker.

   2. In the overall tracker, filter down to the specific facility or facilities you are interested in

   3. Select all the rows showing (including the header row), hit CTRL+C to copy

   4. Press CTRL+N to open a new, blank excel document

   5. In the very first cell (A1), press CTRL+V to paste in all the copied rows. A small box will pop up called paste options, click that then click on the icon that looks like the clipboard with a blank page in front of it and a horizontal line over it (called “Keep source column widths”). This will make all the columns look as they did in the master tracker. 

   6. Click Row 1 (the header row), go to the “Data” tab in Excel, and click the “Filter” button

   7. Rename this tab from “Sheet1” to whatever you want (“Data”, “All Data”, “All Facilities”, the specific facility and it’s WDID number, etc.). Whatever will make sense to you so you know what’s there in case you add more tabs later for other facilities and to differentiate between the other tab we will create in a few steps.

   8. Save this excel as a new file, and remember to include the data when it was created (e.g. “SMARTS Ad Hoc Monitoring Data as of 5-30-24”) so that if you need to update it later with new data, it is easy to remember which data is missing based on whatever data was uploaded to SMARTS after that date (i.e., after 5-30-24) 

   9. Delete all columns that are not necessary (e.g., if you only have one facility in the file, you can delete the WDID, App ID, address, etc.; You can delete the monitoring location description column if they are all blank; you can delete the monitoring location name if there is only 1 discharge point; you can delete the event type column if they are all listed as QSEs). Delete whatever feels like too much information. DO NOT DELETE the Sample ID column because if you ever want to update the data later with new results, you will want to be able to use that column as the reference like when updating the main tracker.

   10. Press CTRL+A to select all cells, and then double click the line between columns A and B (the cursor will change into a plus shape with arrows going left and right). This will resize all the columns based on what you actually have in them (i.e., it’ll make them narrower where they were only wide because of longer entries elsewhere in the main tracker).

   11. Again, this step is optional, but makes reference easier. Because the “reporting year” column only lists the first year (i.e., it lists 2023-2024 as “2023”), this step will convert all the numbers into the correct reporting year. 

       1) Right click on the column to the right of the “Reporting Year” column and click “insert”. This should add a blank column to the right of the “Reporting Year” column.

       2) In the header of this new column, write “Reporting Year” (i.e., copy the header from the original column)

       3) In the first cell over the new column, under the header (i.e., row 2), write the following formula: \=CONCAT(\[original reporting year column\]2,"-",(\[original reporting year column\]2+1)). Because the original reporting year column will depend on what columns you chose to delete earlier, I cannot say the specific column it’ll end up being. But, for example, if the original reporting year column is column A, then the formula would be \=CONCAT(A2,"-",(A2+1))

       4) When the cursor in the bottom right corner of the cell where you wrote that formula turns into the thin black plus sign, double click to automatically fill in the rest of the column. A small box will pop up, click the box and then select “Fill without formatting”. This will fill in the formula but keeps the original formatting

       5) Then click on this new reporting year column, hit CTRL+C to copy, and then right click and under “paste options”, hit the icon that has a little “123” in the corner, which will just paste the answers instead of the formulas. 

       6) Delete the original reporting year column (the one that only displayed the first year), so that you now only have the new reporting year column displaying both years in the range (e.g., 2023-2024).

   12. Because the master tracker likely has some shortcuts regarding which cells have been highlighted (e.g., cells are highlighted red because of the LA River NELs, but this is a Dominguez Channel facility where the NELs do not apply), you need to re-highlight the rows if the receiving water is different than whatever you used for the master tracker.

       1) The easiest way to do this is to press CTRL+A and unhighlight all the cells. Then sort by “parameter” and then “results” from highest to lowest and go through and manually re-highlight the cells as you want

       2) If you want to highlight differently for before and after a limit (e.g., NELs) went into effect, filter for before after that date in the drop down for the “sample date” column and then highlight the cells that you filtered for (E.g., filter for after the NELs went into effect and highlight those cells based on the NELS and filter for before the NELs went into effect and highlight those cells based on the NALs)

   13. VERY IMPORTANT FOR THE DATA VISUALIZATION TO WORK: Excel does not recognize blank cells when it is analyzing data so in the “Result” column drop down, select only “(Blanks)” and then write in the number zero in each cell showing so that there are no blank cells in that column

   14. In the “Data” tab in Excel, hit “Clear” to go back to all the data. 

   15. You are now done with the data and can move on to the visualization.

   16. Press the plus sign on the bottom bar in Excel to create a new sheet and then rename it (I recommend calling it “summary” because that is what this is going to be). 

   17. The rest of the instructions are all in this new sheet.

   18. In the new sheet, go to the “insert” tab in Excel, and click on the “PivotTable” button.

   19. A text box will pop up, click on the first box (“Table/Range”), then click on your sheet with all of your data that you just formatted (whatever you named it). If you did this correctly, you will see that the first box now says the name of this sheet with an exclamation point.

   20. Click on the left-most column and drag all the way to the right-most column to select all columns. In the text box, press “OK”.

   21. This will make a sidebar pop up that says “PivotTable Fields”.

   22. Drag “Parameter” to the top-right box (“Columns”)

   23. Then drag whatever you want to be the rows into the bottom-left box (“Rows”), in the order you want them to appear. For example, if you have more than one facility and they have more than one discharge point each, you might want to do address, then reporting year, then monitoring location. But if you only have one facility with one discharge point, you can just do reporting year and nothing else.

   24. Then drag “results” to the bottom-right box (“Values”). It will automatically start as “Sum of Result” so if you click it, and then select “Value Field Settings”, you can change this to “Average” (for the annual averages) or “Count” (for the number of QSEs sampled). I recommend starting with annual averages.

       1) Also, for a QSEs sampled table, it will count all of them as separate QSEs if there are multiple discharge points or multiple facilities so in all but the bottom of whatever you put in the “Rows” Box of the side bar (e.g., all but “monitoring location”), click on it, select “Value Field Settings” and in the text box that pops up, select “None” for the “subtotals” option, which should be the first thing you see in the textbox.

   25. Then in the “PivotTable Analyze” tab on the top bar of Excel, click on “Options” all the way to left. A text box will pop up. On the first tab (“Layout & Format”), write in what you want empty cells to show (e.g., for annual averages, I write “N/A”, but for QSEs sampled, I write “0”). Then in the second tab (“Totals & Filters”) unclick the top 2 boxes for showing grand totals. Then press OK.

   26. In the drop down for either rows or columns, unclick Blanks so you don’t have those in your table. This now shows the data you need.

   27. If you want to have both annual average and QSEs sampled, the easiest way is to click the first cell in the top-left of the table, which highlights the whole table, and then copy (CTRL+C) and paste (CTRL+V) it to another place in the sheet.

   28. And then you can just manually write in any notes you want to add next to the table (e.g., if there were instantaneous exceedances in that year, if the 4 QSEs that it says were sampled were all in the first half of the year, etc.)