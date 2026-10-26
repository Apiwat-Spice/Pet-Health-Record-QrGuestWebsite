
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

    const { data: visits, error } = await supabase
      .from('visit_records')
      .select(`
        visit_id,
        visit_date,
        visit_time,
        weight,
        symptoms,
        diagnosis,
        treatment,
        note,
        status,
        appointments!inner (
          appointment_id,
          pet_id,
          appointment_date,
          appointment_time,
          title,
          type,
          detail
        )
      `)
      .eq('appointments.pet_id', result.pet.pet_id)
      .in('status', ['Finish', 'Completed'])
      .order('visit_date', { ascending: true })
      .order('visit_time', { ascending: true });

    if (error) throw error;

    const workbook = new ExcelJS.Workbook();
    workbook.creator = 'WebClinic';
    workbook.created = new Date();

    const sheet = workbook.addWorksheet('ประวัติการรักษา', {
      views: [{ state: 'frozen', ySplit: 3 }],
    });

    const pet = result.pet;

    sheet.mergeCells('A1:L1');
    sheet.getCell('A1').value = 'ประวัติการรักษาสัตว์เลี้ยง';
    sheet.getCell('A1').font = {
      name: 'Tahoma',
      size: 16,
      bold: true,
      color: { argb: 'FFFFFFFF' },
    };
    sheet.getCell('A1').fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FF426782' },
    };
    sheet.getCell('A1').alignment = {
      horizontal: 'center',
      vertical: 'middle',
    };
    sheet.getRow(1).height = 32;

    sheet.mergeCells('A2:L2');
    sheet.getCell('A2').value =
      `สัตว์เลี้ยง: ${pet.pet_name} | ประเภท: ${pet.animal_type} | สายพันธุ์: ${pet.species || '-'}`;
    sheet.getCell('A2').font = {
      name: 'Tahoma',
      size: 11,
      bold: true,
    };
    sheet.getCell('A2').alignment = {
      vertical: 'middle',
    };
    sheet.getRow(2).height = 25;

    const headers = [
      'วันที่เข้ารับบริการ',
      'เวลา',
      'วันที่นัดหมาย',
      'ประเภทบริการ',
      'ชื่อบริการ',
      'รูปแบบการเข้ารับบริการ',
      'น้ำหนัก (กก.)',
      'อาการ',
      'การวินิจฉัย',
      'การรักษา',
      'หมายเหตุ',
      'สถานะประวัติ',
    ];

    sheet.addRow(headers);

    const headerRow = sheet.getRow(3);
    headerRow.height = 35;
    headerRow.eachCell((cell) => {
      cell.font = {
        name: 'Tahoma',
        size: 10,
        bold: true,
        color: { argb: 'FFFFFFFF' },
      };
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FF527D9A' },
      };
      cell.alignment = {
        horizontal: 'center',
        vertical: 'middle',
        wrapText: true,
      };
      cell.border = {
        top: { style: 'thin', color: { argb: 'FFD1DCE5' } },
        bottom: { style: 'thin', color: { argb: 'FFD1DCE5' } },
      };
    });

    const serviceColors = {
      'Vaccination': 'FFD9EAD3',
      'Health Checkup': 'FFD9EAF7',
      'Sick Visit / Consultation': 'FFFCE4D6',
      'Spay / Neuter': 'FFE4DFEC',
      'Grooming': 'FFFFF2CC',
      'Bathing': 'FFDDEBF7',
      'Haircut': 'FFF4CCCC',
      'Boarding': 'FFD9EAD3',
      'Grooming / Boarding': 'FFEADCF8',
    };

    const categoryColors = {
      'Walk-in': 'FFE2F0D9',
      'Follow-up': 'FFD9EAF7',
      'Appointment': 'FFD9EAF7',
      'Surgery / Procedure': 'FFF4CCCC',
      'Treatment / Wound Care': 'FFFCE4D6',
    };

    const displayDate = (value) => {
      if (!value) return '-';
      return String(value).slice(0, 10);
    };

    const displayTime = (value) => {
      if (!value) return '-';
      return String(value).slice(0, 5);
    };

    if (!visits || visits.length === 0) {
      sheet.mergeCells('A4:L4');
      sheet.getCell('A4').value = 'ยังไม่มีประวัติการรักษาที่เสร็จสิ้น';
      sheet.getCell('A4').alignment = {
        horizontal: 'center',
        vertical: 'middle',
      };
      sheet.getCell('A4').font = {
        name: 'Tahoma',
        italic: true,
        color: { argb: 'FF808080' },
      };
    } else {
      visits.forEach((visit, index) => {
        const appointment = visit.appointments || {};
        const title = appointment.title || '-';
        const type = appointment.type || '-';

        const row = sheet.addRow([
          displayDate(visit.visit_date),
          displayTime(visit.visit_time),
          displayDate(appointment.appointment_date),
          title === 'Vaccination' ? 'วัคซีน' :
          title === 'Health Checkup' ? 'ตรวจสุขภาพ' :
          title === 'Sick Visit / Consultation' ? 'ตรวจอาการป่วย' :
          title === 'Spay / Neuter' ? 'ทำหมัน' :
          ['Grooming', 'Bathing', 'Haircut', 'Boarding', 'Grooming / Boarding'].includes(title)
            ? 'ดูแลสัตว์เลี้ยง' : 'บริการอื่น',
          title,
          type,
          visit.weight == null ? '-' : Number(visit.weight),
          visit.symptoms || '-',
          visit.diagnosis || '-',
          visit.treatment || '-',
          visit.note || '-',
          visit.status,
        ]);

        row.height = 32;
        row.eachCell((cell) => {
          cell.font = { name: 'Tahoma', size: 10 };
          cell.alignment = {
            vertical: 'top',
            wrapText: true,
          };
          cell.border = {
            bottom: { style: 'thin', color: { argb: 'FFE0E0E0' } },
          };
          if (index % 2 === 1) {
            cell.fill = {
              type: 'pattern',
              pattern: 'solid',
              fgColor: { argb: 'FFF5F8FA' },
            };
          }
        });

        const titleColor = serviceColors[title] || 'FFE0E0E0';
        row.getCell(5).fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: titleColor },
        };
        row.getCell(5).font = {
          name: 'Tahoma',
          size: 10,
          bold: true,
        };

        const typeColor = categoryColors[type] || 'FFE0E0E0';
        row.getCell(6).fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: typeColor },
        };

        row.getCell(1).alignment = {
          horizontal: 'center',
          vertical: 'top',
        };
        row.getCell(2).alignment = {
          horizontal: 'center',
          vertical: 'top',
        };
        row.getCell(3).alignment = {
          horizontal: 'center',
          vertical: 'top',
        };
      });
    }

    sheet.columns = [
      { key: 'visit_date', width: 18 },
      { key: 'visit_time', width: 12 },
      { key: 'appointment_date', width: 18 },
      { key: 'service_category', width: 22 },
      { key: 'service_title', width: 30 },
      { key: 'service_type', width: 25 },
      { key: 'weight', width: 15 },
      { key: 'symptoms', width: 35 },
      { key: 'diagnosis', width: 35 },
      { key: 'treatment', width: 35 },
      { key: 'note', width: 30 },
      { key: 'status', width: 18 },
    ];

    sheet.autoFilter = {
      from: 'A3',
      to: 'L3',
    };

    res.setHeader(
      'Content-Type',
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    );
    res.setHeader(
      'Content-Disposition',
      'attachment; filename="pet-treatment-history.xlsx"',
    );

    await workbook.xlsx.write(res);
    res.end();
  } catch (error) {
    console.error('Export treatment history error:', error.message);
    if (!res.headersSent) {
      res.status(500).send('ไม่สามารถดาวน์โหลดประวัติการรักษาได้');
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
