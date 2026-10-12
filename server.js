
require('dotenv').config();
const express = require('express');
const path = require('path');
const app = express();

app.use(express.urlencoded({ extended: true }));
app.use(express.json());
// app.use(express.static(path.join(__dirname, 'public')));
app.use(express.static(path.join(__dirname, 'src/view/public')));
app.set('views', path.join(__dirname, 'src/view'));
app.set('view engine', 'ejs');

// เพิ่ม Route ของ Guest QR
const guestPetRoutes = require('./src/view/routes/guestPetRoutes');
app.use('/', guestPetRoutes);

// คง Session และ Routes เดิมของ WebClinic ไว้
// app.listen ต้องอยู่หลังจากลงทะเบียน Routes ทั้งหมดแล้ว

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`Web-Guest running at http://localhost:${PORT}`);
});
