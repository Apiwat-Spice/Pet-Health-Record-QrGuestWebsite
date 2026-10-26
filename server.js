
require('dotenv').config();
const express = require('express');
const path = require('path');
const app = express();

app.use(express.urlencoded({ extended: true }));
app.use(express.json());

app.use(express.static(path.join(__dirname, 'src/view/public')));
app.set('views', path.join(__dirname, 'src/view'));
app.set('view engine', 'ejs');

const guestPetRoutes = require('./src/view/routes/guestPetRoutes');
app.use('/', guestPetRoutes);

const PORT = process.env.PORT || 3001;

if (require.main === module) {
  app.listen(PORT, () => {
    console.log(`Web-Guest running at http://localhost:${PORT}`);
  });
}

module.exports = app;
