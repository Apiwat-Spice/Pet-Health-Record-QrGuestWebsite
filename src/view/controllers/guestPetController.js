
const supabase = require('../../../config/supabase');
const QRCode = require('qrcode');
const PDFDocument = require('pdfkit');
const ExcelJS = require('exceljs');
const path = require('path');
const fs = require('fs');

async function getGuestPetData(qrCode) {
  const { data: pet, error: petError } = await supabase
    .from('pets')
    .select('*')
    .eq('qr_code', qrCode)
    .maybeSingle();

  if (petError) throw petError;
  if (!pet) return null;

  const { data: settings, error: settingsError } = await supabase
    .from('ShowQr')
    .select('*')
    .eq('pet_id', pet.pet_id)
    .maybeSingle();

  if (settingsError) throw settingsError;

  const config = settings || {};
  const petNameEnabled = config.Pet_name === true;
  const needsOwnerName = config.Pet_owner_name === true;
  const needsEmail = config.email === true;
  const needsPhone = config.phone === true;

  let ownerName = null;
  let email = null;

  if (needsOwnerName || needsEmail) {
    const { data: owner, error: ownerError } = await supabase
      .from('pet_owners')
      .select('user_id, first_name, last_name')
      .eq('pet_owner_id', pet.pet_owner_id)
      .maybeSingle();

    if (ownerError) throw ownerError;

    if (owner && needsOwnerName) {
      ownerName = `${owner.first_name} ${owner.last_name}`;
    }

    if (owner && needsEmail) {
      const { data: user, error: userError } = await supabase
        .from('users')
        .select('email')
        .eq('user_id', owner.user_id)
        .maybeSingle();

      if (userError) throw userError;
      email = user?.email || null;
    }
  }

  const fields = [
    { label: 'ชื่อสัตว์เลี้ยง', key: 'Pet_name', value: pet.pet_name },
    { label: 'ประเภทสัตว์', key: 'Animal_type', value: pet.animal_type },
    { label: 'สายพันธุ์', key: 'species', value: pet.species },
    { label: 'รายละเอียด', key: 'detail', value: pet.detail },
    { label: 'วันเกิด', key: 'birthdate', value: pet.birthdate },
    { label: 'น้ำหนัก (กก.)', key: 'weight', value: pet.weight },
    { label: 'ชื่อเจ้าของ', key: 'Pet_owner_name', value: ownerName },
    { label: 'อีเมล', key: 'email', value: email },
    { label: 'เบอร์โทรศัพท์', key: 'phone', value: null },
  ].map((field) => {
    const enabled = config[field.key] === true;
    const value =
      field.value === null || field.value === undefined || field.value === ''
        ? '-'
        : String(field.value);

    return {
      key: field.key,
      label: field.label,
      enabled,
      value: enabled ? value : null,
    };
  });

  return {
    pet,
    fields,
    petNameEnabled,
    petDisplayName: petNameEnabled ? pet.pet_name : 'ข้อมูลสัตว์เลี้ยง',
  };
}

function getDisplayValue(field) {
  return field.enabled ? field.value : 'เจ้าของปิดการแสดงผล';
}

exports.showPet = async (req, res) => {
  try {
    const result = await getGuestPetData(req.params.qr_code);

    if (!result) {
      return res.status(404).render('guest/pet', {
        pet: null,
        fields: [],
      });
    }

    res.render('guest/pet', result);
  } catch (error) {
    console.error('Guest pet error:', error.message);
    res.status(500).send('เกิดข้อผิดพลาดในการโหลดข้อมูล');
  }
};

exports.generateQr = async (req, res) => {
  try {
    const { data: pet, error } = await supabase
      .from('pets')
      .select('pet_id')
      .eq('qr_code', req.params.qr_code)
      .maybeSingle();

    if (error) throw error;
    if (!pet) return res.status(404).send('ไม่พบข้อมูลสัตว์เลี้ยง');

    const baseUrl =
      process.env.PUBLIC_BASE_URL || `${req.protocol}://${req.get('host')}`;

    const url = `${baseUrl}/guest/pet/${encodeURIComponent(req.params.qr_code)}`;

    const image = await QRCode.toBuffer(url, {
      type: 'png',
      width: 400,
      margin: 2,
      errorCorrectionLevel: 'M',
    });

    res.type('png').send(image);
  } catch (error) {
    console.error('Generate QR error:', error.message);
    res.status(500).send('ไม่สามารถสร้าง QR Code ได้');
  }
};

exports.downloadExcel = async (req, res) => {
  try {
    const result = await getGuestPetData(req.params.qr_code);

    if (!result) {
      return res.status(404).send('ไม่พบข้อมูลสัตว์เลี้ยง');
    }

    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('Pet Information');

    sheet.columns = [
      { header: 'หัวข้อ', key: 'label', width: 25 },
      { header: 'ข้อมูล', key: 'value', width: 40 },
    ];

    sheet.addRow({ label: 'สัตว์เลี้ยง', value: result.petDisplayName });

    result.fields.forEach((field) => {
      sheet.addRow({
        label: field.label,
        value: getDisplayValue(field),
      });
    });

    sheet.getRow(1).font = { bold: true };
    sheet.getRow(2).font = { bold: true };
    sheet.views = [{ state: 'frozen', ySplit: 2 }];

    res.setHeader(
      'Content-Type',
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    );
    res.setHeader(
      'Content-Disposition',
      'attachment; filename="pet-information.xlsx"',
    );

    await workbook.xlsx.write(res);
    res.end();
  } catch (error) {
    console.error('Export Excel error:', error.message);
    if (!res.headersSent) {
      res.status(500).send('ไม่สามารถดาวน์โหลด Excel ได้');
    }
  }
};

exports.downloadPdf = async (req, res) => {
  try {
    const result = await getGuestPetData(req.params.qr_code);

    if (!result) {
      return res.status(404).send('ไม่พบข้อมูลสัตว์เลี้ยง');
    }

    const fontPath = path.join(
      __dirname,
      '../public/css/fonts/NotoSansThai-Regular.ttf',
    );

    if (!fs.existsSync(fontPath)) {
      console.error('ไม่พบไฟล์ฟอนต์:', fontPath);
      return res.status(500).send('ไม่พบไฟล์ฟอนต์ภาษาไทย');
    }

    const doc = new PDFDocument({
      size: 'A4',
      margin: 45,
      info: { Title: 'Pet Information' },
    });

    doc.on('error', (error) => {
      console.error('PDF stream error:', error.message);
      if (!res.destroyed) res.destroy();
    });

    res.setHeader('Content-Type', 'application/pdf');
    res.setHeader(
      'Content-Disposition',
      'attachment; filename="pet-information.pdf"',
    );

    doc.pipe(res);
    doc.font(fontPath);

    const leftX = doc.page.margins.left;
    const rightMargin = doc.page.width - doc.page.margins.right;
    const rightX = leftX + 165;
    const rightWidth = rightMargin - rightX;
    const rowGap = 10;

function drawSectionTitle(title, align = 'center', indent = 0) {
  doc.font(fontPath)
    .fontSize(20)
    .text(title, doc.page.margins.left + indent, doc.y, {
      width: doc.page.width - doc.page.margins.left - doc.page.margins.right - indent,
      align,
    });

  doc.moveDown(0.5);

  const lineY = doc.y;
  doc.moveTo(doc.page.margins.left, lineY)
    .lineTo(doc.page.width - doc.page.margins.right, lineY)
    .lineWidth(1)
    .strokeColor('#8EB4D0')
    .stroke();

  doc.moveDown(1);
}
    function drawField(field) {
      const label = `${field.label} :`;
      const value = getDisplayValue(field);
      const startY = doc.y;
      const labelWidth = 155;

      doc.font(fontPath).fontSize(11).text(label, leftX, startY, {
        width: labelWidth,
      });

      const labelHeight = doc.heightOfString(label, {
        width: labelWidth,
        font: fontPath,
        fontSize: 11,
      });

      doc.font(fontPath).fontSize(12).text(value, rightX, startY, {
        width: rightWidth,
        lineGap: 2,
      });

      const valueHeight = doc.heightOfString(value, {
        width: rightWidth,
        font: fontPath,
        fontSize: 12,
        lineGap: 2,
      });

      doc.y = startY + Math.max(labelHeight, valueHeight) + rowGap;

      if (doc.y > doc.page.height - doc.page.margins.bottom - 30) {
        doc.addPage();
      }
    }

    const petFields = result.fields.filter((field) =>
      ['Pet_name', 'Animal_type', 'species', 'detail', 'birthdate', 'weight']
        .includes(field.key),
    );

    const ownerFields = result.fields.filter((field) =>
      ['Pet_owner_name', 'email', 'phone'].includes(field.key),
    );

    drawSectionTitle('ข้อมูลสัตว์เลี้ยง', 'center');

    doc.font(fontPath).fontSize(16).text(result.petDisplayName);
    doc.moveDown(1);

    petFields.forEach(drawField);

    if (ownerFields.some((field) => field.enabled)) {
      if (doc.y > doc.page.height - doc.page.margins.bottom - 100) {
        doc.addPage();
      }

      doc.moveDown(1);
      drawSectionTitle('ข้อมูลเจ้าของสัตว์เลี้ยง', 'center');

      ownerFields
        .filter((field) => field.enabled)
        .forEach(drawField);
    }

    doc.end();
  } catch (error) {
    console.error('Export PDF error:', error.message);

    if (!res.headersSent && !res.destroyed) {
      return res.status(500).send('ไม่สามารถดาวน์โหลด PDF ได้');
    }

    if (!res.destroyed) res.destroy();
  }
};
