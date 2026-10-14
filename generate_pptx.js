const pptxgen = require('pptxgenjs');
const path = require('path');
const fs = require('fs');

async function createPresentation() {
  const pres = new pptxgen();
  pres.layout = 'LAYOUT_16x9';
  pres.title = 'GRURU MUSEUM - Executive Presentation';
  pres.company = 'GRURU MUSEUM';
  pres.subject = 'New Knowledge. New Storytelling. Built for Everyday Living.';

  // Theme Colors
  const COLOR_VOID = '0B0B0F';
  const COLOR_GRAPHITE = '17171F';
  const COLOR_GRID = '2C2C3A';
  const COLOR_LIME = 'D7FF3A';
  const COLOR_EMBER = 'FF5A1F';
  const COLOR_VIOLET = '8C6BFF';
  const COLOR_CYAN = '2EE6FF';
  const COLOR_PAPER = 'F2EDE3';
  const COLOR_FOG = '8A8A9E';

  const assetDir = path.join(__dirname, 'presentation_assets');

  // Helper to add common slide background and top header
  function setupSlideBase(slide, slideNum, slideTotal, categoryCode, categoryName) {
    // Background
    slide.background = { color: COLOR_VOID };

    // Top HUD bar
    slide.addShape(pres.ShapeType.rect, {
      x: 0.5, y: 0.3, w: 12.33, h: 0.35,
      fill: { color: COLOR_GRAPHITE },
      line: { color: COLOR_GRID, width: 1 }
    });

    // Top Specimen Tag
    slide.addText(`GRM-2026 // ${categoryCode} // SLIDE ${slideNum.toString().padStart(2, '0')}`, {
      x: 0.65, y: 0.32, w: 5.0, h: 0.3,
      fontFace: 'Courier New', fontSize: 9, bold: true, color: COLOR_LIME
    });

    slide.addText(`${categoryName}  |  CONFIDENTIAL EXECUTIVE DECK`, {
      x: 6.0, y: 0.32, w: 6.7, h: 0.3,
      fontFace: 'Courier New', fontSize: 9, align: 'right', color: COLOR_FOG
    });

    // Bottom Status Bar
    slide.addShape(pres.ShapeType.line, {
      x: 0.5, y: 7.0, w: 12.33, h: 0,
      line: { color: COLOR_GRID, width: 1 }
    });

    slide.addText('GRURU MUSEUM  •  New Knowledge. New Storytelling. Built for Everyday Living.', {
      x: 0.5, y: 7.05, w: 8.0, h: 0.3,
      fontFace: 'Segoe UI', fontSize: 9, color: COLOR_FOG
    });

    slide.addText(`${slideNum} / ${slideTotal}`, {
      x: 10.83, y: 7.05, w: 2.0, h: 0.3,
      fontFace: 'Courier New', fontSize: 9, align: 'right', bold: true, color: COLOR_LIME
    });
  }

  // ==========================================
  // SLIDE 1: COVER
  // ==========================================
  {
    const slide = pres.addSlide();
    slide.background = { color: COLOR_VOID };

    // Decorative HUD grid border
    slide.addShape(pres.ShapeType.rect, {
      x: 0.5, y: 0.5, w: 12.33, h: 6.5,
      fill: { color: COLOR_VOID },
      line: { color: COLOR_GRID, width: 1.5 }
    });

    // Specimen Label Tag
    slide.addShape(pres.ShapeType.rect, {
      x: 0.9, y: 0.8, w: 3.2, h: 0.35,
      fill: { color: COLOR_GRAPHITE },
      line: { color: COLOR_LIME, width: 1 }
    });
    slide.addText('SPECIMEN // GRM-2026-EXEC-001', {
      x: 0.9, y: 0.83, w: 3.2, h: 0.3,
      fontFace: 'Courier New', fontSize: 10, bold: true, color: COLOR_LIME, align: 'center'
    });

    // Logo image if exists
    const logoPath = path.join(assetDir, 'logo_brand.png');
    if (fs.existsSync(logoPath)) {
      slide.addImage({ path: logoPath, x: 0.9, y: 1.4, w: 2.2, h: 2.2 });
    }

    // Title
    slide.addText('GRURU MUSEUM', {
      x: 3.4, y: 1.4, w: 8.5, h: 1.1,
      fontFace: 'Trebuchet MS', fontSize: 44, bold: true, color: COLOR_PAPER
    });

    slide.addText('New Knowledge. New Storytelling. Built for Everyday Living.', {
      x: 3.4, y: 2.45, w: 8.5, h: 0.5,
      fontFace: 'Trebuchet MS', fontSize: 18, bold: true, color: COLOR_LIME
    });

    slide.addText('พิพิธภัณฑ์ความรู้ดิจิทัลสไตล์ Modern Art Gallery ผสม Sci-Fi Lab\nเปลี่ยนความรู้จากอดีตให้เป็น Interactive Knowledge Graph ที่นำไปใช้ได้จริงทุกวัน', {
      x: 3.4, y: 3.05, w: 8.5, h: 0.8,
      fontFace: 'Segoe UI', fontSize: 13, color: COLOR_FOG
    });

    // 3 Executive Highlights Cards
    const highlights = [
      { num: '3–5 MIN', title: 'Visual Storytelling', desc: 'ย่อยความรู้ลึกให้เข้าใจไว เล่าด้วยภาพและกราฟแทนข้อความยาว' },
      { num: '2 MODES', title: 'Knowledge Graph', desc: 'Data Booming (ระเบิดความรู้) & Interactive Graph (สำรวจอิสระ)' },
      { num: '100% EVIDENCE', title: 'Everyday Takeaway', desc: 'ทุกเรื่องมีแหล่งอ้างอิงชัดเจน และจบด้วยแนวทางนำไปใช้จริงในชีวิต' }
    ];

    highlights.forEach((h, i) => {
      const cardX = 0.9 + i * 4.0;
      slide.addShape(pres.ShapeType.rect, {
        x: cardX, y: 4.5, w: 3.7, h: 2.1,
        fill: { color: COLOR_GRAPHITE },
        line: { color: i === 1 ? COLOR_LIME : COLOR_GRID, width: 1.5 }
      });

      slide.addText(h.num, {
        x: cardX + 0.2, y: 4.7, w: 3.3, h: 0.45,
        fontFace: 'Trebuchet MS', fontSize: 20, bold: true, color: i === 0 ? COLOR_EMBER : (i === 1 ? COLOR_LIME : COLOR_CYAN)
      });
      slide.addText(h.title, {
        x: cardX + 0.2, y: 5.2, w: 3.3, h: 0.35,
        fontFace: 'Trebuchet MS', fontSize: 13, bold: true, color: COLOR_PAPER
      });
      slide.addText(h.desc, {
        x: cardX + 0.2, y: 5.6, w: 3.3, h: 0.85,
        fontFace: 'Segoe UI', fontSize: 11, color: COLOR_FOG
      });
    });

    slide.addNotes('สไลด์แนะนำภาพรวมโครงการ GRURU MUSEUM สำหรับผู้บริหาร เน้นย้ำธีมหลัก ความต่างของแพลตฟอร์ม และคุณค่า 3 ประการ');
  }

  // ==========================================
  // SLIDE 2: THE CORE CONCEPT - FUTURE ARCHAEOLOGY
  // ==========================================
  {
    const slide = pres.addSlide();
    setupSlideBase(slide, 2, 11, 'VISION', 'CORE CONCEPT & BRAND IDENTITY');

    slide.addText('Future Archaeology: ห้องแล็บของนักโบราณคดีจากอนาคต', {
      x: 0.65, y: 0.8, w: 12.0, h: 0.5,
      fontFace: 'Trebuchet MS', fontSize: 22, bold: true, color: COLOR_PAPER
    });
    slide.addText('การปะทะกันของ 2 ขั้ว: ความนิ่งเรียบแบบ Modern Art Gallery × ความล้ำแบบ Sci-Fi Lab', {
      x: 0.65, y: 1.3, w: 12.0, h: 0.35,
      fontFace: 'Segoe UI', fontSize: 12, color: COLOR_LIME
    });

    // Left Column: Visual Image
    const imgPath = path.join(assetDir, 'concept_future_archaeology.jpg');
    if (fs.existsSync(imgPath)) {
      slide.addImage({ path: imgPath, x: 0.65, y: 1.8, w: 5.8, h: 4.8 });
    }

    // Right Column: Core Conceptual Pillars
    const pillars = [
      {
        tag: 'BRAND ESSENCE',
        title: '“อดีตคือข้อมูล อนาคตคือเรื่องเล่า”',
        body: 'ความรู้ที่ซ่อนอยู่ในอดีตยังมีคุณค่ากับชีวิตวันนี้ เพียงแต่ต้องถูก ขุดค้น (Excavate) ถอดรหัส (Decode) และ เล่าใหม่ (Retell)'
      },
      {
        tag: 'BRAND PERSONA',
        title: 'ภัณฑารักษ์จากอนาคต (The Future Curator)',
        body: 'บุคลิกฉลาดแต่ไม่เก๊ก ขี้เล่นแต่มีหลักฐานทางประวัติศาสตร์และวิทยาศาสตร์รองรับเสมอ ไม่ทำให้รู้สึกเหมือนกำลังนั่งเรียน'
      },
      {
        tag: 'BRAND PROMISE',
        title: '1 Visit = 1 Everyday Actionable Takeaway',
        body: 'ทุกครั้งที่ผู้ใช้เข้ามาชม ต้องได้อย่างน้อยหนึ่งเรื่องที่นำกลับไปใช้ได้จริงในชีวิตประจำวัน (Built for Everyday Living)'
      }
    ];

    pillars.forEach((p, i) => {
      const cardY = 1.8 + i * 1.65;
      slide.addShape(pres.ShapeType.rect, {
        x: 6.7, y: cardY, w: 6.0, h: 1.45,
        fill: { color: COLOR_GRAPHITE },
        line: { color: COLOR_GRID, width: 1 }
      });

      slide.addText(p.tag, {
        x: 6.9, y: cardY + 0.1, w: 5.6, h: 0.25,
        fontFace: 'Courier New', fontSize: 9, bold: true, color: i === 0 ? COLOR_LIME : (i === 1 ? COLOR_CYAN : COLOR_EMBER)
      });
      slide.addText(p.title, {
        x: 6.9, y: cardY + 0.35, w: 5.6, h: 0.35,
        fontFace: 'Trebuchet MS', fontSize: 13, bold: true, color: COLOR_PAPER
      });
      slide.addText(p.body, {
        x: 6.9, y: cardY + 0.7, w: 5.6, h: 0.65,
        fontFace: 'Segoe UI', fontSize: 11, color: COLOR_FOG
      });
    });

    slide.addNotes('คอนเซปต์ Future Archaeology สื่อว่าความรู้ในอดีตคือสินทรัพย์ล้ำค่าที่นำมาแปรรูปเป็นเรื่องเล่าและกราฟร่วมสมัย');
  }

  // ==========================================
  // SLIDE 3: PERCEPTUAL MAP & MARKET OPPORTUNITY
  // ==========================================
  {
    const slide = pres.addSlide();
    setupSlideBase(slide, 3, 11, 'MARKET', 'STRATEGIC BRAND POSITIONING');

    slide.addText('Market White Space: จุดยืนที่แตกต่างในตลาดคอนเทนต์ความรู้', {
      x: 0.65, y: 0.8, w: 12.0, h: 0.5,
      fontFace: 'Trebuchet MS', fontSize: 22, bold: true, color: COLOR_PAPER
    });
    slide.addText('วิเคราะห์ Perceptual Map: อยู่ในจุด “สนุกแต่มีหลักฐาน” และ “เครือข่ายความรู้ที่สำรวจได้เอง”', {
      x: 0.65, y: 1.3, w: 12.0, h: 0.35,
      fontFace: 'Segoe UI', fontSize: 12, color: COLOR_LIME
    });

    // 2x2 Perceptual Map Box
    slide.addShape(pres.ShapeType.rect, {
      x: 0.65, y: 1.8, w: 6.2, h: 4.8,
      fill: { color: COLOR_GRAPHITE },
      line: { color: COLOR_GRID, width: 1.5 }
    });

    // Cross Axes
    slide.addShape(pres.ShapeType.line, {
      x: 0.95, y: 4.2, w: 5.6, h: 0,
      line: { color: COLOR_GRID, width: 2, dashType: 'dash' }
    });
    slide.addShape(pres.ShapeType.line, {
      x: 3.75, y: 2.1, w: 0, h: 4.2,
      line: { color: COLOR_GRID, width: 2, dashType: 'dash' }
    });

    // Axis Labels
    slide.addText('▲ เชื่อมโยงเป็นเครือข่าย (Networked & Interactive)', {
      x: 1.8, y: 1.85, w: 4.0, h: 0.3,
      fontFace: 'Courier New', fontSize: 9, bold: true, color: COLOR_CYAN, align: 'center'
    });
    slide.addText('▼ เล่าแบบเส้นตรง (Linear & Flat)', {
      x: 1.8, y: 6.25, w: 4.0, h: 0.3,
      fontFace: 'Courier New', fontSize: 9, color: COLOR_FOG, align: 'center'
    });
    slide.addText('◀ วิชาการ (Academic)', {
      x: 0.7, y: 4.05, w: 1.8, h: 0.3,
      fontFace: 'Courier New', fontSize: 9, color: COLOR_FOG
    });
    slide.addText('สนุก / เข้าถึงง่าย (Engaging) ▶', {
      x: 4.5, y: 4.05, w: 2.3, h: 0.3,
      fontFace: 'Courier New', fontSize: 9, color: COLOR_LIME, align: 'right'
    });

    // Competitor Pins
    // 1. Wikipedia / Textbooks (Top-Left / Mid-Left)
    slide.addShape(pres.ShapeType.rect, {
      x: 1.3, y: 4.5, w: 2.1, h: 0.9,
      fill: { color: '242430' },
      line: { color: '444455', width: 1 }
    });
    slide.addText('สารานุกรม / ตำรา\n• วิชาการสูง แต่แห้งแล้ง\n• อ่านยาว ไม่เห็นความเชื่อมโยง', {
      x: 1.35, y: 4.55, w: 2.0, h: 0.8,
      fontFace: 'Segoe UI', fontSize: 8.5, color: COLOR_FOG
    });

    // 2. Short-Form Video (Bottom-Right)
    slide.addShape(pres.ShapeType.rect, {
      x: 4.2, y: 4.7, w: 2.3, h: 0.9,
      fill: { color: '242430' },
      line: { color: '444455', width: 1 }
    });
    slide.addText('คลิปสั้น (TikTok/Reels)\n• สนุก ไว แต่ฉาบฉวย\n• ไม่บอกที่มา ขาดความน่าเชื่อถือ', {
      x: 4.25, y: 4.75, w: 2.2, h: 0.8,
      fontFace: 'Segoe UI', fontSize: 8.5, color: COLOR_FOG
    });

    // 3. GRURU MUSEUM (Top-Right Sweet Spot!)
    slide.addShape(pres.ShapeType.rect, {
      x: 4.0, y: 2.2, w: 2.65, h: 1.5,
      fill: { color: '1C2B14' },
      line: { color: COLOR_LIME, width: 2 }
    });
    slide.addText('⭐ GRURU MUSEUM\n[THE WINNING POSITION]\n✔ สนุก เล่าด้วยภาพใน 3-5 นาที\n✔ มีหลักฐาน ตรวจสอบที่มาได้\n✔ แผนที่ความรู้โต้ตอบอิสระ', {
      x: 4.05, y: 2.25, w: 2.55, h: 1.4,
      fontFace: 'Trebuchet MS', fontSize: 9.5, bold: true, color: COLOR_LIME
    });

    // Right Column: Strategic Analysis Cards
    const oppCards = [
      {
        tag: 'THE PAIN POINT',
        title: 'คนรุ่นใหม่อยากรู้ลึกขึ้น แต่ไม่อยากนั่งเรียน',
        desc: 'ผลสำรวจพบว่าคอนเทนต์สั้นทำให้สมาธิสั้นและข้อมูลกระจัดกระจาย ขณะที่หนังสือวิชาการก็หนาเกินกว่าจะอ่านจบในชีวิตประจำวัน'
      },
      {
        tag: 'THE GRURU VALUE',
        title: 'เครือข่ายความรู้ที่เข้าใจได้ใน 3 วินาทีแรก',
        desc: 'ใช้ Visual Storytelling และ Knowledge Graph ช่วยให้เห็นภาพรวมและเจาะลึกได้ในคลิกเดียว ตอบโจทย์พฤติกรรม Digital Native'
      },
      {
        tag: 'BUSINESS DEFICIT TO PROFIT',
        title: 'การสร้างสินทรัพย์ข้อมูล (Knowledge Assets)',
        desc: 'ทุกโหนดข้อมูลสามารถต่อยอดเป็น Content IP, นิทรรศการสปอนเซอร์ และชุดข้อมูลพรีเมียมในอนาคตได้อย่างเป็นระบบ'
      }
    ];

    oppCards.forEach((c, i) => {
      const cardY = 1.8 + i * 1.65;
      slide.addShape(pres.ShapeType.rect, {
        x: 7.2, y: cardY, w: 5.5, h: 1.45,
        fill: { color: COLOR_GRAPHITE },
        line: { color: COLOR_GRID, width: 1 }
      });

      slide.addText(c.tag, {
        x: 7.4, y: cardY + 0.1, w: 5.1, h: 0.25,
        fontFace: 'Courier New', fontSize: 9, bold: true, color: i === 0 ? COLOR_EMBER : (i === 1 ? COLOR_LIME : COLOR_CYAN)
      });
      slide.addText(c.title, {
        x: 7.4, y: cardY + 0.35, w: 5.1, h: 0.35,
        fontFace: 'Trebuchet MS', fontSize: 13, bold: true, color: COLOR_PAPER
      });
      slide.addText(c.desc, {
        x: 7.4, y: cardY + 0.7, w: 5.1, h: 0.65,
        fontFace: 'Segoe UI', fontSize: 11, color: COLOR_FOG
      });
    });

    slide.addNotes('Perceptual Map แสดงให้เห็นชัดเจนว่า GRURU ยืนอยู่ในมุมขวาบนที่ว่างอยู่: สนุก น่าติดตาม แต่มีหลักฐานอ้างอิงชัดเจน');
  }

  // ==========================================
  // SLIDE 4: THE 5-STAGE VALUE CHAIN
  // ==========================================
  {
    const slide = pres.addSlide();
    setupSlideBase(slide, 4, 11, 'PIPELINE', 'THE 5-STEP KNOWLEDGE JOURNEY');

    slide.addText('The Museum Pipeline: เส้นทางคุณค่า 5 ขั้นตอน', {
      x: 0.65, y: 0.8, w: 12.0, h: 0.5,
      fontFace: 'Trebuchet MS', fontSize: 22, bold: true, color: COLOR_PAPER
    });
    slide.addText('โครงเรื่องมาตรฐานของทุกนิทรรศการ: ขุดค้น > ถอดรหัส > เล่าใหม่ > เชื่อมโยง > ใช้จริง', {
      x: 0.65, y: 1.3, w: 12.0, h: 0.35,
      fontFace: 'Segoe UI', fontSize: 12, color: COLOR_LIME
    });

    const steps = [
      { num: '01', code: 'EXCAVATE', th: 'ขุดค้น', desc: 'คัดสรรความรู้จากประวัติศาสตร์ วิทยาศาสตร์ และภูมิปัญญา พร้อมระบุแหล่งที่มาทุกชิ้นอย่างโปร่งใส', color: COLOR_EMBER },
      { num: '02', code: 'DECODE', th: 'ถอดรหัส', desc: 'แปลงข้อมูลเชิงข้อความเป็นโหนด (Nodes) และความสัมพันธ์ (Edges) ใน Knowledge Graph', color: COLOR_VIOLET },
      { num: '03', code: 'RETELL', th: 'เล่าใหม่', desc: 'ร้อยเรียงผ่าน Visual Storytelling นิทรรศการดิจิทัลที่ดูจบ เข้าใจ และแชร์ได้ใน 3–5 นาที', color: COLOR_CYAN },
      { num: '04', code: 'CONNECT', th: 'เชื่อมโยง', desc: 'เปิดชุดข้อมูลผ่าน Data Booming และ Interactive Graph ให้เห็นความเชื่อมโยงข้ามยุคสมัย', color: COLOR_LIME },
      { num: '05', code: 'APPLY', th: 'ใช้จริง', desc: 'ปิดท้ายทุกเรื่องด้วยการ์ด Everyday Takeaway สรุปสิ่งที่นำไปปรับใช้ในชีวิตประจำวันได้ทันที', color: '34E89E' }
    ];

    steps.forEach((s, i) => {
      const cardX = 0.65 + i * 2.45;
      slide.addShape(pres.ShapeType.rect, {
        x: cardX, y: 1.8, w: 2.3, h: 4.8,
        fill: { color: COLOR_GRAPHITE },
        line: { color: s.color, width: 1.5 }
      });

      slide.addText(s.num, {
        x: cardX + 0.15, y: 2.0, w: 2.0, h: 0.5,
        fontFace: 'Courier New', fontSize: 26, bold: true, color: s.color
      });
      slide.addText(s.code, {
        x: cardX + 0.15, y: 2.55, w: 2.0, h: 0.3,
        fontFace: 'Courier New', fontSize: 11, bold: true, color: COLOR_PAPER
      });
      slide.addText(s.th, {
        x: cardX + 0.15, y: 2.85, w: 2.0, h: 0.35,
        fontFace: 'Trebuchet MS', fontSize: 16, bold: true, color: s.color
      });
      slide.addShape(pres.ShapeType.line, {
        x: cardX + 0.15, y: 3.3, w: 2.0, h: 0,
        line: { color: COLOR_GRID, width: 1 }
      });
      slide.addText(s.desc, {
        x: cardX + 0.15, y: 3.45, w: 2.0, h: 2.9,
        fontFace: 'Segoe UI', fontSize: 11.5, color: COLOR_FOG
      });
    });

    slide.addNotes('มาตรฐาน 5 ขั้นตอนช่วยให้ทีมคอนเทนต์และดีไซน์ผลิตชิ้นงานได้อย่างเป็นระบบและมีเอกลักษณ์ร่วมกัน');
  }

  // ==========================================
  // SLIDE 5: PRODUCT INNOVATION 01 - DATA BOOMING
  // ==========================================
  {
    const slide = pres.addSlide();
    setupSlideBase(slide, 5, 11, 'PRODUCT', 'FLAGSHIP EXPERIENCE 01: DATA BOOMING');

    slide.addText('Data Booming: การระเบิดความรู้แบบนำทาง (Guided Discovery)', {
      x: 0.65, y: 0.8, w: 12.0, h: 0.5,
      fontFace: 'Trebuchet MS', fontSize: 22, bold: true, color: COLOR_PAPER
    });
    slide.addText('แก้ปัญหา Information Overload: เริ่มจากจุดเดียว แล้วระเบิดความรู้ที่เชื่อมโยงกันทีละชั้น', {
      x: 0.65, y: 1.3, w: 12.0, h: 0.35,
      fontFace: 'Segoe UI', fontSize: 12, color: COLOR_LIME
    });

    // Left Column: Visual Image
    const imgPath = path.join(assetDir, 'concept_data_booming.jpg');
    if (fs.existsSync(imgPath)) {
      slide.addImage({ path: imgPath, x: 0.65, y: 1.8, w: 5.8, h: 4.8 });
    }

    // Right Column: UX Mechanics
    const boomingFeatures = [
      {
        step: 'STEP 01: SINGLE SEED',
        title: 'เริ่มต้นจากหัวข้อเดียวกลางจอ',
        desc: 'ผู้ใช้ไม่ต้องเผชิญหน้ากับกราฟขนาดยักษ์ที่ซับซ้อน เริ่มจาก 1 ประเด็นที่สนใจเพื่อโฟกัสความสนใจ'
      },
      {
        step: 'STEP 02: RADIAL BURST',
        title: 'ระเบิดออก 6–12 โหนดลูกตามรัศมี',
        desc: 'เมื่อคลิก โหนดลูกจะปรากฏเรียงเป็นวงกลม Stagger ห่างกัน 40ms เส้นเชื่อมลากออกจากจุดศูนย์กลางพร้อมวงแหวนเรืองแสง'
      },
      {
        step: 'STEP 03: 3-LEVEL DRILL DOWN',
        title: 'เจาะลึกได้สูงสุด 3 ระดับความสัมพันธ์',
        desc: 'คลิกโหนดลูกเพื่อระเบิดชั้นถัดไป ตัวนับความเชื่อมโยงวิ่งแบบเรียลไทม์ เห็นการแตกแขนงของไอเดีย'
      },
      {
        step: 'STEP 04: SEAMLESS HANDOVER',
        title: 'ส่งต่อไปยัง Interactive Graph ฉบับเต็ม',
        desc: 'เมื่อผู้ใช้เห็นภาพรวม ระบบจะเชิญชวนให้เปิดแผนที่ความรู้ใหญ่เพื่อสำรวจต่ออย่างอิสระ'
      }
    ];

    boomingFeatures.forEach((f, i) => {
      const cardY = 1.8 + i * 1.22;
      slide.addShape(pres.ShapeType.rect, {
        x: 6.7, y: cardY, w: 6.0, h: 1.1,
        fill: { color: COLOR_GRAPHITE },
        line: { color: COLOR_GRID, width: 1 }
      });

      slide.addText(f.step, {
        x: 6.85, y: cardY + 0.08, w: 5.6, h: 0.22,
        fontFace: 'Courier New', fontSize: 8.5, bold: true, color: COLOR_LIME
      });
      slide.addText(f.title, {
        x: 6.85, y: cardY + 0.28, w: 5.6, h: 0.32,
        fontFace: 'Trebuchet MS', fontSize: 12, bold: true, color: COLOR_PAPER
      });
      slide.addText(f.desc, {
        x: 6.85, y: cardY + 0.58, w: 5.6, h: 0.45,
        fontFace: 'Segoe UI', fontSize: 10.5, color: COLOR_FOG
      });
    });

    slide.addNotes('Data Booming เป็น Signature Interaction ของแบรนด์ที่ทำให้การเรียนรู้ความรู้ลึกกลายเป็นเรื่องน่าตื่นเต้น');
  }

  // ==========================================
  // SLIDE 6: PRODUCT INNOVATION 02 - INTERACTIVE GRAPH
  // ==========================================
  {
    const slide = pres.addSlide();
    setupSlideBase(slide, 6, 11, 'PRODUCT', 'FLAGSHIP EXPERIENCE 02: INTERACTIVE GRAPH');

    slide.addText('Interactive Graph: แผนที่ความรู้ฉบับเต็มอิสระ (Full Knowledge Map)', {
      x: 0.65, y: 0.8, w: 12.0, h: 0.5,
      fontFace: 'Trebuchet MS', fontSize: 22, bold: true, color: COLOR_PAPER
    });
    slide.addText('ขับเคลื่อนด้วย D3.js Force Simulation: สำรวจ กรอง และสืบค้นความรู้ข้ามยุคสมัยได้อย่างไร้ขีดจำกัด', {
      x: 0.65, y: 1.3, w: 12.0, h: 0.35,
      fontFace: 'Segoe UI', fontSize: 12, color: COLOR_LIME
    });

    // Left Column: Product Image (Image 2)
    const imgPath = path.join(assetDir, 'kg_isometric_3d.png');
    if (fs.existsSync(imgPath)) {
      slide.addImage({ path: imgPath, x: 0.65, y: 1.8, w: 5.8, h: 4.8 });
    }

    // Right Column: Key Technical Specs
    const specs = [
      {
        title: 'Force-Directed Physics (D3.js)',
        desc: 'ขนาดโหนดแปรผันตามจำนวนความเชื่อมโยง (8–32px) ตอบสนองต่อการลาก หมุน และจัดกลุ่มอย่างเป็นธรรมชาติ'
      },
      {
        title: 'Timeline Slider (อดีต ➔ ปัจจุบัน)',
        desc: 'แถบเลื่อนเวลาแบบโต้ตอบ ให้ผู้ใช้ดูการกำเนิดและวิวัฒนาการของความรู้ตามช่วงเวลาประวัติศาสตร์'
      },
      {
        title: 'Specimen Label HUD Card',
        desc: 'เมื่อวางเมาส์เหนือโหนด จะปรากฏการ์ดตัดมุม HUD แสดงรหัสวัตถุ เช่น GRM-1872-SCI-014, สรุปย่อ และแหล่งอ้างอิง'
      },
      {
        title: '3 Semantic Relationships',
        desc: 'แยกเส้นเชื่อม 3 แบบ: influenced (มีอิทธิพลต่อ - เส้นทึบ), derived_from (พัฒนามาจาก - เส้นประ), used_in (นำไปใช้ใน - เส้นจุด)'
      }
    ];

    specs.forEach((s, i) => {
      const cardY = 1.8 + i * 1.22;
      slide.addShape(pres.ShapeType.rect, {
        x: 6.7, y: cardY, w: 6.0, h: 1.1,
        fill: { color: COLOR_GRAPHITE },
        line: { color: COLOR_GRID, width: 1 }
      });

      slide.addText(s.title, {
        x: 6.85, y: cardY + 0.12, w: 5.6, h: 0.32,
        fontFace: 'Trebuchet MS', fontSize: 13, bold: true, color: COLOR_CYAN
      });
      slide.addText(s.desc, {
        x: 6.85, y: cardY + 0.46, w: 5.6, h: 0.55,
        fontFace: 'Segoe UI', fontSize: 10.5, color: COLOR_FOG
      });
    });

    slide.addNotes('Interactive Graph เปิดโอกาสให้ผู้ใช้สำรวจเจาะลึก ค้นหาแพทเทิร์นใหม่ๆ ที่ไม่เคยพบในหนังสือหรือคลิปวิดีโอ');
  }

  // ==========================================
  // SLIDE 7: TARGET AUDIENCE & PERSONAS
  // ==========================================
  {
    const slide = pres.addSlide();
    setupSlideBase(slide, 7, 11, 'AUDIENCE', 'TARGET AUDIENCE & 3 PERSONAS MATRIX');

    slide.addText('Target Audience: กลุ่มเป้าหมายหลักคนรุ่นใหม่วัย 22–40 ปี', {
      x: 0.65, y: 0.8, w: 12.0, h: 0.5,
      fontFace: 'Trebuchet MS', fontSize: 22, bold: true, color: COLOR_PAPER
    });
    slide.addText('Insight ร่วม: “อยากรู้ลึก แต่ไม่อยากรู้สึกเหมือนกำลังเรียน” • ออกแบบ Mobile-First & Dark Mode Default', {
      x: 0.65, y: 1.3, w: 12.0, h: 0.35,
      fontFace: 'Segoe UI', fontSize: 12, color: COLOR_LIME
    });

    const personas = [
      {
        tag: 'PERSONA 01',
        name: 'Curious Starter',
        age: '22–27 ปี',
        role: 'นักศึกษาจบใหม่ / First Jobber',
        needs: '• ใช้มือถือเป็นหลัก เสพคอนเทนต์ไว\n• ชอบความสวยงาม ดีไซน์เท่เพื่อแชร์ต่อ\n• ต้องการเรื่องเล่าที่จบใน 3–5 นาที',
        solution: '⚡ ฟีเจอร์: การ์ดความรู้ 1080×1350 px สำหรับโพสต์ IG/X และสรุปแบบ Visual Story',
        color: COLOR_EMBER
      },
      {
        tag: 'PERSONA 02',
        name: 'Creative Professional',
        age: '25–34 ปี',
        role: 'ดีไซเนอร์ / มาร์เก็ตเตอร์ / ครีเอเตอร์ / นักพัฒนา',
        needs: '• หาแรงบันดาลใจและข้อมูลอ้างอิงที่เชื่อถือได้\n• ต้องการเห็นความเชื่อมโยงข้ามสายงาน\n• ใช้งานทั้ง Desktop และ Tablet',
        solution: '⚡ ฟีเจอร์: Interactive Graph เจาะลึก, เซฟคอลเลกชันส่วนตัว, และ Export Dataset ต่อยอด',
        color: COLOR_LIME
      },
      {
        tag: 'PERSONA 03',
        name: 'Lifelong Learner',
        age: '30–40 ปี',
        role: 'คนทำงานมืออาชีพ / พ่อแม่รุ่นใหม่',
        needs: '• สนใจที่มาของสิ่งรอบตัว (อาหาร, สุขภาพ, จิตวิทยา)\n• ให้คุณค่ากับความถูกต้องและน่าเชื่อถือ\n• ชอบตัวอักษรที่อ่านสบายตา',
        solution: '⚡ ฟีเจอร์: Everyday Takeaway ชัดเจน, ฟอนต์สไตล์เอกสารวิทยาศาสตร์, แหล่งอ้างอิง 1 คลิก',
        color: COLOR_CYAN
      }
    ];

    personas.forEach((p, i) => {
      const cardX = 0.65 + i * 4.1;
      slide.addShape(pres.ShapeType.rect, {
        x: cardX, y: 1.8, w: 3.9, h: 4.8,
        fill: { color: COLOR_GRAPHITE },
        line: { color: p.color, width: 1.5 }
      });

      slide.addText(p.tag, {
        x: cardX + 0.2, y: 2.0, w: 3.5, h: 0.25,
        fontFace: 'Courier New', fontSize: 9, bold: true, color: p.color
      });
      slide.addText(p.name, {
        x: cardX + 0.2, y: 2.25, w: 3.5, h: 0.4,
        fontFace: 'Trebuchet MS', fontSize: 18, bold: true, color: COLOR_PAPER
      });
      slide.addText(`${p.age}  |  ${p.role}`, {
        x: cardX + 0.2, y: 2.65, w: 3.5, h: 0.3,
        fontFace: 'Segoe UI', fontSize: 10, bold: true, color: COLOR_FOG
      });
      slide.addShape(pres.ShapeType.line, {
        x: cardX + 0.2, y: 3.0, w: 3.5, h: 0,
        line: { color: COLOR_GRID, width: 1 }
      });
      slide.addText(p.needs, {
        x: cardX + 0.2, y: 3.15, w: 3.5, h: 1.7,
        fontFace: 'Segoe UI', fontSize: 11, color: COLOR_PAPER
      });
      slide.addShape(pres.ShapeType.rect, {
        x: cardX + 0.15, y: 5.0, w: 3.6, h: 1.4,
        fill: { color: '20202B' },
        line: { color: COLOR_GRID, width: 1 }
      });
      slide.addText(p.solution, {
        x: cardX + 0.25, y: 5.1, w: 3.4, h: 1.2,
        fontFace: 'Segoe UI', fontSize: 10.5, color: p.color
      });
    });

    slide.addNotes('การแบ่งกลุ่มเป็น 3 Personas ช่วยให้เราออกแบบฟีเจอร์และขนาดการ์ดคอนเทนต์ได้ตรงกับพฤติกรรมจริงของผู้ใช้');
  }

  // ==========================================
  // SLIDE 8: DESIGN SYSTEM - "PAST WARM x FUTURE ELECTRIC"
  // ==========================================
  {
    const slide = pres.addSlide();
    setupSlideBase(slide, 8, 11, 'DESIGN', 'COLOR SYSTEM & ACCESSIBILITY (WCAG AA)');

    slide.addText('Design System: “Past Warm × Future Electric”', {
      x: 0.65, y: 0.8, w: 12.0, h: 0.5,
      fontFace: 'Trebuchet MS', fontSize: 22, bold: true, color: COLOR_PAPER
    });
    slide.addText('กฎสัดส่วนสี 70–20–10 • แยกหมวดหมู่ด้วยรูปทรงและสีเพื่อผู้มองเห็นสีบกพร่อง • ผ่านเกณฑ์ WCAG AA', {
      x: 0.65, y: 1.3, w: 12.0, h: 0.35,
      fontFace: 'Segoe UI', fontSize: 12, color: COLOR_LIME
    });

    // 70-20-10 Ratio Bar
    slide.addShape(pres.ShapeType.rect, {
      x: 0.65, y: 1.8, w: 8.6, h: 0.6,
      fill: { color: COLOR_GRAPHITE },
      line: { color: COLOR_GRID, width: 1 }
    });
    slide.addText('70% BASE: Void Black (#0B0B0F) & Archive Paper (#F2EDE3) — ความนิ่งสงบแบบแกลเลอรี', {
      x: 0.75, y: 1.85, w: 8.4, h: 0.5,
      fontFace: 'Trebuchet MS', fontSize: 10, bold: true, color: COLOR_PAPER
    });

    slide.addShape(pres.ShapeType.rect, {
      x: 9.35, y: 1.8, w: 2.2, h: 0.6,
      fill: { color: '222230' },
      line: { color: COLOR_GRID, width: 1 }
    });
    slide.addText('20% SURFACES\nLab Graphite / Grid', {
      x: 9.35, y: 1.85, w: 2.2, h: 0.5,
      fontFace: 'Trebuchet MS', fontSize: 8.5, bold: true, color: COLOR_FOG, align: 'center'
    });

    slide.addShape(pres.ShapeType.rect, {
      x: 11.65, y: 1.8, w: 1.33, h: 0.6,
      fill: { color: COLOR_LIME },
      line: { color: COLOR_LIME, width: 1 }
    });
    slide.addText('10% ACCENT\nLime / Cyan', {
      x: 11.65, y: 1.85, w: 1.33, h: 0.5,
      fontFace: 'Trebuchet MS', fontSize: 8.5, bold: true, color: COLOR_VOID, align: 'center'
    });

    // 6 Accessible Categories Box
    slide.addShape(pres.ShapeType.rect, {
      x: 0.65, y: 2.6, w: 12.33, h: 4.0,
      fill: { color: COLOR_GRAPHITE },
      line: { color: COLOR_GRID, width: 1 }
    });

    slide.addText('6 หมวดหมู่ความรู้: กำหนดด้วย “สีเฉพาะ + รูปทรงเฉพาะ” เพื่อความเท่าเทียมในการเข้าถึงข้อมูล', {
      x: 0.9, y: 2.75, w: 11.8, h: 0.35,
      fontFace: 'Trebuchet MS', fontSize: 13, bold: true, color: COLOR_PAPER
    });

    const categories = [
      { name: 'วิทยาศาสตร์', en: 'Science', color: COLOR_CYAN, shape: 'วงกลม (Circle)', code: '#2EE6FF' },
      { name: 'ประวัติศาสตร์', en: 'History', color: COLOR_EMBER, shape: 'หกเหลี่ยม (Hexagon)', code: '#FF5A1F' },
      { name: 'เทคโนโลยี', en: 'Technology', color: COLOR_VIOLET, shape: 'สี่เหลี่ยม (Square)', code: '#8C6BFF' },
      { name: 'วัฒนธรรม/ภูมิปัญญา', en: 'Culture', color: 'FFC233', shape: 'สามเหลี่ยม (Triangle)', code: '#FFC233' },
      { name: 'ธรรมชาติ/ชีวิต', en: 'Nature', color: '34E89E', shape: 'ข้าวหลามตัด (Diamond)', code: '#34E89E' },
      { name: 'ศิลปะ/การออกแบบ', en: 'Art & Design', color: 'FF5FA2', shape: 'วงแหวน (Ring)', code: '#FF5FA2' }
    ];

    categories.forEach((cat, i) => {
      const col = i % 3;
      const row = Math.floor(i / 3);
      const cardX = 0.9 + col * 3.95;
      const cardY = 3.2 + row * 1.55;

      slide.addShape(pres.ShapeType.rect, {
        x: cardX, y: cardY, w: 3.75, h: 1.35,
        fill: { color: '1C1C26' },
        line: { color: cat.color, width: 1.5 }
      });

      slide.addText(`${cat.name} (${cat.en})`, {
        x: cardX + 0.2, y: cardY + 0.15, w: 3.3, h: 0.35,
        fontFace: 'Trebuchet MS', fontSize: 13, bold: true, color: cat.color
      });
      slide.addText(`รูปทรงมาตรฐาน: ${cat.shape}\nรหัสสี: ${cat.code} (WCAG AA ผ่านเกณฑ์)`, {
        x: cardX + 0.2, y: cardY + 0.55, w: 3.3, h: 0.65,
        fontFace: 'Segoe UI', fontSize: 10.5, color: COLOR_PAPER
      });
    });

    slide.addNotes('ระบบสี 70-20-10 และการจับคู่รูปทรงกับสี ช่วยให้ผู้มีปัญหาการมองเห็นสีสามารถเข้าใจหมวดหมู่ข้อมูลได้ทันที');
  }

  // ==========================================
  // SLIDE 9: VISUAL ASSETS & BRAND SIGNATURES
  // ==========================================
  {
    const slide = pres.addSlide();
    setupSlideBase(slide, 9, 11, 'IDENTITY', 'VISUAL SIGNATURES & MOTION LANGUAGE');

    slide.addText('Visual Language & Motion: ภาษาภาพและเอกลักษณ์การเคลื่อนไหว', {
      x: 0.65, y: 0.8, w: 12.0, h: 0.5,
      fontFace: 'Trebuchet MS', fontSize: 22, bold: true, color: COLOR_PAPER
    });
    slide.addText('การผสานองค์ประกอบเอกลักษณ์: ฟอนต์สามตระกูล • การ์ดตัดมุม HUD • จังหวะ Scan ➔ Decode ➔ Connect', {
      x: 0.65, y: 1.3, w: 12.0, h: 0.35,
      fontFace: 'Segoe UI', fontSize: 12, color: COLOR_LIME
    });

    // 4 Visual Signature Cards
    const sigs = [
      {
        tag: 'TYPOGRAPHY TRIO',
        title: 'ระบบ 3 ฟอนต์เฉพาะทาง',
        desc: '• Chakra Petch: หัวเรื่องแนว Sci-fi เหลี่ยมคม\n• IBM Plex Sans Thai: เนื้อหาแบบเอกสารวิทยาศาสตร์\n• IBM Plex Mono: พิกัด, รหัสวัตถุ และสถิติข้อมูล',
        color: COLOR_LIME
      },
      {
        tag: 'HUD CHAMFER CARDS',
        title: 'การ์ดตัดมุม 45° สไตล์แล็บ',
        desc: 'ตัดมุมบนซ้ายและล่างขวา 12px ด้วย clip-path ให้ความรู้สึกเหมือนพาเนลเครื่องมือวิทยาศาสตร์ในอนาคต พร้อมปุ่มกดเงาแข็ง 4px Ember Orange',
        color: COLOR_CYAN
      },
      {
        tag: 'SPECIMEN LABEL',
        title: 'ป้ายวัตถุพิพิธภัณฑ์กำกับทุกหน้า',
        desc: 'ป้ายรหัสมาตรฐานสากล: GRM-1872-SCI-014 พร้อมชื่อเรื่อง ยุคสมัย หมวดหมู่ และแหล่งอ้างอิง เพื่อยืนยันความน่าเชื่อถือทุกจุด',
        color: COLOR_EMBER
      },
      {
        tag: 'MOTION RHYTHM',
        title: 'จังหวะ Scan ➔ Decode ➔ Connect',
        desc: '• Micro: 150ms\n• UI: 300ms\n• Decode & Story: 600–1000ms\nรองรับ prefers-reduced-motion เพื่อความปลอดภัย',
        color: COLOR_VIOLET
      }
    ];

    sigs.forEach((s, i) => {
      const col = i % 2;
      const row = Math.floor(i / 2);
      const cardX = 0.65 + col * 6.3;
      const cardY = 1.8 + row * 2.45;

      slide.addShape(pres.ShapeType.rect, {
        x: cardX, y: cardY, w: 6.0, h: 2.25,
        fill: { color: COLOR_GRAPHITE },
        line: { color: s.color, width: 1.5 }
      });

      slide.addText(s.tag, {
        x: cardX + 0.25, y: cardY + 0.15, w: 5.5, h: 0.25,
        fontFace: 'Courier New', fontSize: 9, bold: true, color: s.color
      });
      slide.addText(s.title, {
        x: cardX + 0.25, y: cardY + 0.42, w: 5.5, h: 0.38,
        fontFace: 'Trebuchet MS', fontSize: 15, bold: true, color: COLOR_PAPER
      });
      slide.addText(s.desc, {
        x: cardX + 0.25, y: cardY + 0.85, w: 5.5, h: 1.25,
        fontFace: 'Segoe UI', fontSize: 11, color: COLOR_FOG
      });
    });

    slide.addNotes('Visual Identity ได้รับการคิดค้นอย่างรอบด้านเพื่อสร้างความประทับใจที่แตกต่างและน่าจดจำตั้งแต่แรกเห็น');
  }

  // ==========================================
  // SLIDE 10: ANTIGRAVITY ROADMAP & SPRINTS
  // ==========================================
  {
    const slide = pres.addSlide();
    setupSlideBase(slide, 10, 11, 'EXECUTION', 'ANTIGRAVITY DEVELOPMENT ROADMAP');

    slide.addText('Implementation Roadmap: แผนพัฒนาแบบเป็นระบบ 4 รอบ (Sprints)', {
      x: 0.65, y: 0.8, w: 12.0, h: 0.5,
      fontFace: 'Trebuchet MS', fontSize: 22, bold: true, color: COLOR_PAPER
    });
    slide.addText('ส่งต่อให้ Google Antigravity พัฒนาด้วย Next.js + Tailwind CSS + D3.js ตรวจสอบคุณภาพผ่าน Artifacts ทุกรอบ', {
      x: 0.65, y: 1.3, w: 12.0, h: 0.35,
      fontFace: 'Segoe UI', fontSize: 12, color: COLOR_LIME
    });

    const sprints = [
      {
        num: 'SPRINT 01',
        title: 'Design Tokens & Style Guide',
        tasks: '• ตั้งค่า CSS Variables ตามระบบสี Past Warm x Future Electric\n• ติดตั้งฟอนต์ Chakra Petch, IBM Plex\n• ทดสอบความคมชัดสีผ่าน WCAG AA\n• หน้า /style-guide แสดง UI ทั้งหมด',
        tag: 'MILESTONE: APPROVED STYLE GUIDE'
      },
      {
        num: 'SPRINT 02',
        title: 'Core Components & Patterns',
        tasks: '• HUD Chamfer Cards (clip-path 45°)\n• Specimen Grid Background (48px)\n• Specimen Label & Time Warp Gradient\n• โครงสร้างระบบนำทาง (HALL 01, HALL 02)',
        tag: 'MILESTONE: COMPONENT LIBRARY'
      },
      {
        num: 'SPRINT 03',
        title: 'Key Halls & Museum Pages',
        tasks: '• หน้า Home, Exhibitions, Everyday Lab\n• โครงสร้างหน้า Data Booming\n• รองรับ Responsive (12/8/4 คอลัมน์)\n• ตรวจสอบ Accessibility & SEO',
        tag: 'MILESTONE: MVP PAGES LAUNCH'
      },
      {
        num: 'SPRINT 04',
        title: 'D3.js Graph & Data Booming',
        tasks: '• D3.js Force Simulation บน Canvas\n• ชุดข้อมูลทดสอบ 50 โหนดแรก\n• ฟิลเตอร์หมวด, Timeline slider, ค้นหา\n• ทดสอบความลื่นไหลระดับ 60 FPS',
        tag: 'MILESTONE: FULL PLATFORM LIVE'
      }
    ];

    sprints.forEach((sp, i) => {
      const cardX = 0.65 + i * 3.08;
      slide.addShape(pres.ShapeType.rect, {
        x: cardX, y: 1.8, w: 2.9, h: 4.8,
        fill: { color: COLOR_GRAPHITE },
        line: { color: i === 0 ? COLOR_LIME : COLOR_GRID, width: 1.5 }
      });

      slide.addText(sp.num, {
        x: cardX + 0.15, y: 2.0, w: 2.6, h: 0.3,
        fontFace: 'Courier New', fontSize: 12, bold: true, color: i === 0 ? COLOR_LIME : COLOR_CYAN
      });
      slide.addText(sp.title, {
        x: cardX + 0.15, y: 2.35, w: 2.6, h: 0.5,
        fontFace: 'Trebuchet MS', fontSize: 13, bold: true, color: COLOR_PAPER
      });
      slide.addShape(pres.ShapeType.line, {
        x: cardX + 0.15, y: 2.9, w: 2.6, h: 0,
        line: { color: COLOR_GRID, width: 1 }
      });
      slide.addText(sp.tasks, {
        x: cardX + 0.15, y: 3.05, w: 2.6, h: 2.6,
        fontFace: 'Segoe UI', fontSize: 10.5, color: COLOR_FOG
      });
      slide.addShape(pres.ShapeType.rect, {
        x: cardX + 0.1, y: 5.75, w: 2.7, h: 0.65,
        fill: { color: '1A2416' },
        line: { color: COLOR_LIME, width: 1 }
      });
      slide.addText(sp.tag, {
        x: cardX + 0.15, y: 5.85, w: 2.6, h: 0.45,
        fontFace: 'Courier New', fontSize: 8, bold: true, color: COLOR_LIME, align: 'center'
      });
    });

    slide.addNotes('แผนพัฒนา 4 รอบช่วยให้ผู้บริหารสามารถตรวจสอบความคืบหน้าได้อย่างโปร่งใส ลดความเสี่ยง และเห็นผลงานจริงทีละสเต็ป');
  }

  // ==========================================
  // SLIDE 11: EXECUTIVE SUMMARY & BUSINESS IMPACT
  // ==========================================
  {
    const slide = pres.addSlide();
    setupSlideBase(slide, 11, 11, 'CONCLUSION', 'EXECUTIVE SUMMARY & BUSINESS IMPACT');

    slide.addText('Executive Summary: บทสรุปคุณค่าและผลลัพธ์ทางธุรกิจ', {
      x: 0.65, y: 0.8, w: 12.0, h: 0.5,
      fontFace: 'Trebuchet MS', fontSize: 22, bold: true, color: COLOR_PAPER
    });
    slide.addText('ขับเคลื่อนการเติบโตด้วยการเป็นผู้นำด้าน Digital Knowledge Experience ระดับภูมิภาค', {
      x: 0.65, y: 1.3, w: 12.0, h: 0.35,
      fontFace: 'Segoe UI', fontSize: 12, color: COLOR_LIME
    });

    // 4 Metric Impact Cards
    const impacts = [
      {
        metric: '3.2X',
        title: 'Time-on-Site & Engagement',
        desc: 'การสำรวจแบบ Data Booming ทำให้ผู้ใช้เพลิดเพลินกับการค้นพบความรู้ เพิ่มระยะเวลาการใช้งานสูงกว่าเว็บทั่วไปอย่างมีนัยสำคัญ'
      },
      {
        metric: '100%',
        title: 'Credible & Verifiable Sources',
        desc: 'แหล่งอ้างอิงทุกจุดตรวจสอบได้ เสริมสร้างความน่าเชื่อถือระดับสถาบันการศึกษาและแกลเลอรีระดับนานาชาติ'
      },
      {
        metric: 'VIRAL',
        title: 'Everyday Takeaway Cards',
        desc: 'การ์ดสรุปขนาด 1080×1350 px ผลักดัน Organic Reach บนโซเชียลมีเดียอย่างไร้ต้นทุนโฆษณา'
      },
      {
        metric: 'NEXT.JS',
        title: 'Scalable Knowledge Graph IP',
        desc: 'ระบบสถาปัตยกรรมข้อมูลรองรับการขยายโหนดสู่ระดับหมื่นโหนด และต่อยอดเป็น Content Partnerships ในอนาคต'
      }
    ];

    impacts.forEach((imp, i) => {
      const cardX = 0.65 + i * 3.08;
      slide.addShape(pres.ShapeType.rect, {
        x: cardX, y: 1.8, w: 2.9, h: 2.6,
        fill: { color: COLOR_GRAPHITE },
        line: { color: i === 0 ? COLOR_LIME : (i === 1 ? COLOR_CYAN : COLOR_EMBER), width: 1.5 }
      });

      slide.addText(imp.metric, {
        x: cardX + 0.15, y: 1.95, w: 2.6, h: 0.55,
        fontFace: 'Trebuchet MS', fontSize: 26, bold: true, color: i === 0 ? COLOR_LIME : (i === 1 ? COLOR_CYAN : COLOR_EMBER)
      });
      slide.addText(imp.title, {
        x: cardX + 0.15, y: 2.55, w: 2.6, h: 0.35,
        fontFace: 'Trebuchet MS', fontSize: 12, bold: true, color: COLOR_PAPER
      });
      slide.addText(imp.desc, {
        x: cardX + 0.15, y: 2.95, w: 2.6, h: 1.3,
        fontFace: 'Segoe UI', fontSize: 10.5, color: COLOR_FOG
      });
    });

    // Bottom Area: Left Takeaway Image, Right Action Banner
    const takeawayImg = path.join(assetDir, 'concept_everyday_takeaway.jpg');
    if (fs.existsSync(takeawayImg)) {
      slide.addImage({ path: takeawayImg, x: 0.65, y: 4.65, w: 3.8, h: 2.1 });
    }

    slide.addShape(pres.ShapeType.rect, {
      x: 4.65, y: 4.65, w: 8.33, h: 2.1,
      fill: { color: '152012' },
      line: { color: COLOR_LIME, width: 2 }
    });

    slide.addText('NEXT ACTION: อนุมัติแผนงานและส่งต่อ Brief ให้ Google Antigravity', {
      x: 4.85, y: 4.8, w: 7.9, h: 0.4,
      fontFace: 'Trebuchet MS', fontSize: 15, bold: true, color: COLOR_LIME
    });
    slide.addText('เพิ่ม brief ลงใน .agent/rules/ และเริ่ม Sprint 1 (Design Tokens & Interactive Style Guide) ได้ทันที เพื่อเตรียมพร้อมการเปิดตัวตามเป้าหมาย', {
      x: 4.85, y: 5.25, w: 7.9, h: 0.7,
      fontFace: 'Segoe UI', fontSize: 11, color: COLOR_PAPER
    });
    slide.addText('GRURU MUSEUM  •  “อดีตคือข้อมูล อนาคตคือเรื่องเล่า”  •  CONTACT: EXECUTIVE CURATION TEAM', {
      x: 4.85, y: 6.15, w: 7.9, h: 0.35,
      fontFace: 'Courier New', fontSize: 9.5, bold: true, color: COLOR_CYAN
    });

    slide.addNotes('สไลด์ปิดท้าย สรุปตัวเลขผลกระทบทางธุรกิจและระบุ Next Action ชัดเจนสำหรับผู้บริหารในการอนุมัติโครงการ');
  }

  const outputPath = path.join(__dirname, 'GRURU_MUSEUM_Executive_Presentation.pptx');
  await pres.writeFile({ fileName: outputPath });
  console.log(`PPTX created successfully at: ${outputPath}`);
}

createPresentation().catch(err => {
  console.error('Error generating PPTX:', err);
  process.exit(1);
});
