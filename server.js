const express = require('express');
const bodyParser = require('body-parser');
const session = require('express-session');
const path = require('path');
const xlsx = require('xlsx');

const app = express();
const PORT = 3000;

// Middleware
app.use(bodyParser.urlencoded({ extended: true }));
app.use(session({
    secret: 'your-secret-key',
    resave: false,
    saveUninitialized: true,
}));

// Serve static files (CSS, images, etc.)
app.use(express.static(path.join(__dirname, 'public')));

// Set EJS as the template engine
app.set('view engine', 'ejs');
app.set('views', path.join(__dirname, 'views'));

// Load Excel Data
const excelFile = path.join(__dirname, 'database.xlsx');

// Function to read data from an Excel sheet
function readExcel(sheetName) {
    const workbook = xlsx.readFile(excelFile);
    const sheet = workbook.Sheets[sheetName];
    return xlsx.utils.sheet_to_json(sheet);
}

// Function to write data to an Excel sheet
function writeExcel(sheetName, data) {
    const workbook = xlsx.readFile(excelFile);
    const worksheet = xlsx.utils.json_to_sheet(data);
    workbook.Sheets[sheetName] = worksheet;
    xlsx.writeFile(workbook, excelFile);
}

// Routes
// Login Page
app.get('/', (req, res) => {
    res.render('login');
});

// Login Form Submission
app.post('/login', (req, res) => {
    const { username, password } = req.body;
    const users = readExcel('Users');

    const user = users.find(u => u.Username === username && u.Password === password);

    if (user) {
        req.session.user = user;
        res.redirect('/dashboard');
    } else {
        res.send('Invalid credentials! <a href="/">Go back</a>');
    }
});

// Dashboard
app.get('/dashboard', (req, res) => {
    if (!req.session.user) {
        return res.redirect('/');
    }

    const slots = readExcel('Slots').filter(slot => slot.Username === req.session.user.Username);
    res.render('dashboard', { user: req.session.user, slots });
});

// Book Slot Page
app.get('/book', (req, res) => {
    if (!req.session.user) {
        return res.redirect('/');
    }
    res.render('book', { successMessage: null }); // Pass successMessage as null initially
});

// Book Slot Form Submission
app.post('/book', (req, res) => {
    const { date, time } = req.body;
    const user = req.session.user;

    const slots = readExcel('Slots');
    slots.push({ Username: user.Username, Date: date, Time: time, Status: 'Booked' });
    writeExcel('Slots', slots);

    // Pass successMessage to the book.ejs view
    res.render('book', { successMessage: 'Slot booked successfully!' });
});

// Logout
app.get('/logout', (req, res) => {
    req.session.destroy();
    res.redirect('/');
});

// Start Server
app.listen(PORT, () => {
    console.log(`Server running on http://localhost:${PORT}`);
});
