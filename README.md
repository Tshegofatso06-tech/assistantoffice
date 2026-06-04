# OfficeAssistant

AI-powered office assistant that turns ANY Excel file into instant reports, emails, notes, and data insights. Built for busy office workers, analysts, and students. An AI web app that saves office workers time by turning Excel data + quick tasks into instant results.
Most office workers waste 3+  hours every day turning Excel spreadsheets into reports and emails, but Office-Assistant fixes that. It’s an AI data analyst that reads ANY Excel file - sales, lab data, farms, surveys - and instantly gives you summaries, professional emails, meeting notes, plus smart analysis like trends, key metrics, and sales insights with one click.

Why this works 
1. Problem first → Shows you understand real pain 
2. ANY Excel → Flexibility = bigger market
3. 9 tools + AI analysis → Sounds advanced, not basic
4. 3 seconds → Speed sells

### What it does
Stop wasting 3+ hours writing reports from spreadsheets. Upload any Excel file and get AI results in 5 seconds.

4 Core AI Tools:
1. Smart Summary - Reads any Excel and writes a 3-sentence summary. Works with sales, lab data, farms, students, inventory, etc
2. Draft Email - Creates professional emails based on your actual data 
3. Write Notes - Generates meeting notes/bullet points from your spreadsheet
4. Send Instructions - Writes clear team instructions based on the data
5. Analyses- Provide analysis like trends, key metrics, and sales insights

AI Data Analysis Suite
5. Key Metrics - Auto-calculates Total, Average, Min, Max, Count based on your columns
6. Trends & Patterns - Finds growth, drops, seasonal patterns, outliers over time
7. Comparison - Auto-detects categories and shows Top 3 vs Bottom 3 
8. Scientific Summary - Analyzes lab/experimental data: what was measured, sample size, key findings
9. Sales Insights - If sales data detected: gives 3 actionable business tips like best product, slow periods

### Tech Stack
- Frontend*: React + Tailwind CSS + Vite
- AI: Lovable AI + OpenAI API for data analysis
- File Parsing: xlsx / read-excel-file for .xlsx files
- Deployment: Lovable + GitHub

### Key Features
- Works with ANY Excel*: Not limited to sales. Reads farms, students, scientific data, surveys, inventory
- *Smart Date Detection*: If "Date" column exists → This Month/Last Month/Week filters work. If not → summarizes all data
- Table Preview: Shows first 10 rows so you know AI sees the right data
- Mobile-friendly: Works on phone + computer

### How to Run Locally
git clone https://github.com/Tshegofatso06-tech/assistantoffice.git
cd assistantoffice
npm install
npm run dev

### How to Use
1. Open the app link
2. Upload any `.xlsx` file with data
3. Click the tool you need: Summary, Email, Notes, Instructions, or Analysis buttons
4. Copy the AI result and use it

### Test Files Included
Upload these to test all features:
- `sales_full.xlsx` - Tests Sales Insights + Trends + Metrics
- `lab_data.xlsx` - Tests Scientific Summary + Trends 
- `farms.xlsx` - Tests Comparison button
- `survey.xlsx` - Tests no-date-column handling

### Problem I Solved
Built v1 with hardcoded sales columns. Debugged file parsing issues so it now adapts to any columns automatically. Added data analysis because office workers need insights, not just summaries.

### Next Steps
- Add user login + save history
- Connect Google Sheets directly  
- Export AI results to PDF/Word

Challenges
Debugged generic file parsing

How was it tested
I upload a farm spreadsheet, click 'Scientific Summary', and AI tells me cattle counts and top farms in 3 seconds. Same web works for sales data too.

### Responsible AI Practices
- **Human-in-the-loop**: AI output should be reviewed before sending emails or making decisions
- **Data privacy**: Files are processed in browser memory only, not stored on servers
- **Transparency**: Users see data preview before AI analysis to verify correctness
- **Limitations**: AI may misinterpret data. Always validate totals and insights

### Sample Prompts Used - Prompt Engineering Proof
1. Draft Email Prompt
"Based on columns, Write professional email with Subject line. 4-5 sentences. 
If Date column exists mention date range. If numeric columns exist include totals. 
Adapt tone for client/manager/team based on data type."

2. Trends & Patterns Prompt:
"Find 3 trends in this data. If Date column exists check growth over time. 
Find highest/lowest values, correlations between columns, and statistical outliers. 
Explain in simple business language."

3. Weekly Plan Prompt:
"Create structured weekly task plan from this data. Prioritize by urgency/importance. 
Suggest time optimization: batch similar tasks, delegate low-value work. 
Format as Monday to Friday daily plan."
