const express = require('express');
const fs = require('fs').promises; // Using the modern 'promises' version of fs for cleaner code
const path = require('path');
const XLSX = require('xlsx');

const app = express();
const PORT = process.env.PORT || 3000;

// --- Configuration ---
// Define paths to important directories and files
const PUBLIC_DIR = path.join(__dirname, 'public');
const PROJECT_REPORTS_DIR = path.join(PUBLIC_DIR, 'ProjectReports2');
const EXCEL_FILE = path.join(__dirname, 'contacts.xlsx');


// --- Middleware ---
// Securely serve all static files (HTML, CSS, images, PDFs) ONLY from the 'public' directory
app.use(express.static(PUBLIC_DIR));
// Use the built-in Express middleware to parse JSON request bodies (replaces body-parser)
app.use(express.json());


// --- API Endpoints ---

// API: List all categories (subfolders) in ProjectReports2
app.get('/api/project-report-categories', async (req, res) => {
    try {
        const files = await fs.readdir(PROJECT_REPORTS_DIR, { withFileTypes: true });
        const categories = files
            .filter(file => file.isDirectory())
            .map(file => file.name)
            .sort((a, b) => a.localeCompare(b)); // Sort alphabetically
        res.json(categories);
    } catch (error) {
        console.error('Error reading categories directory:', error);
        res.status(500).json({ error: 'Unable to list categories' });
    }
});

// API: List all PDFs in a given category
app.get('/api/project-reports', async (req, res) => {
    const { category } = req.query;
    if (!category) {
        return res.status(400).json({ error: 'Category is required.' });
    }

    try {
        const categoryDir = path.join(PROJECT_REPORTS_DIR, category);
        const files = await fs.readdir(categoryDir);
        
        const pdfs = files
            .filter(file => file.toLowerCase().endsWith('.pdf'))
            .map(file => ({
                // Clean up the name for display (remove .pdf extension and underscores)
                name: path.basename(file, '.pdf').replace(/_/g, ' '),
                // Provide the correct public URL for the file so it can be downloaded
                file: `/ProjectReports2/${encodeURIComponent(category)}/${encodeURIComponent(file)}`
            }));
        res.json(pdfs);
    } catch (error) {
        console.error('Error listing files in category:', error);
        res.status(500).json({ error: 'Unable to list files for the specified category' });
    }
});

// API: Save contact details to the Excel file
app.post('/api/contact', async (req, res) => {
    const { name, email, mobile } = req.body;
    if (!name || !email || !mobile) {
        return res.status(400).json({ error: 'All fields are required.' });
    }

    try {
        let contacts = [];
        // Check if the Excel file already exists
        try {
            await fs.access(EXCEL_FILE); // Check for file existence asynchronously
            const workbook = XLSX.readFile(EXCEL_FILE);
            const sheetName = workbook.SheetNames[0];
            contacts = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);
        } catch {
            // File doesn't exist, we will create it.
            console.log('Contacts file not found, creating a new one.');
        }
        
        // Add the new contact with a timestamp
        contacts.push({ name, email, mobile, date: new Date().toISOString() });

        // Write the updated data back to the Excel file
        const worksheet = XLSX.utils.json_to_sheet(contacts);
        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, worksheet, 'Contacts');
        XLSX.writeFile(workbook, EXCEL_FILE);

        res.status(201).json({ success: true, message: 'Contact saved.' });
    } catch (error) {
        console.error('Error saving contact to Excel:', error);
        res.status(500).json({ error: 'Failed to save contact.' });
    }
});


// --- Start Server ---
app.listen(PORT, () => {
    console.log(`Server is running at http://localhost:${PORT}`);
});