
const express = require('express');
const router = express.Router();
const guestPetController = require('../controllers/guestPetController');

router.get('/guest/pet/:qr_code', guestPetController.showPet);
router.get('/guest/pet/:qr_code/qr.png', guestPetController.generateQr);
router.get('/guest/pet/:qr_code/download/pdf', guestPetController.downloadPdf);
router.get('/guest/pet/:qr_code/download/excel', guestPetController.downloadExcel);

module.exports = router;
