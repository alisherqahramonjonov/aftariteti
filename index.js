const path = require('path');
require('dotenv').config({ path: path.resolve(__dirname, '.env') });
const TelegramBot = require('node-telegram-bot-api');
const PptxGenJS = require('pptxgenjs');
const fs = require('fs');
const Groq = require('groq-sdk');
const express = require('express');
const axios = require('axios');
const app = express();

// 🖼️ Rasm yuklash funksiyasi
async function fetchImageBase64(prompt) {
  try {
    const url = `https://image.pollinations.ai/prompt/${encodeURIComponent(prompt)}?width=1024&height=768&nologo=true`;
    const response = await axios.get(url, { responseType: 'arraybuffer', timeout: 15000 });
    const base64 = Buffer.from(response.data, 'binary').toString('base64');
    return `image/jpeg;base64,${base64}`;
  } catch(e) {
    console.error("Rasm tortishda xatolik:", e.message);
    return null;
  }
}

// 🤖 AI orqali rasm uchun batafsil tavsif (prompt) yaratish
async function generateImagePrompt(topic) {
  try {
    const res = await groq.chat.completions.create({
      model: "llama-3.3-70b-versatile",
      messages: [
        { role: "system", content: "You are an expert at creating visual prompts for AI image generators. Create a professional, cinematic, high-quality background illustration description for a presentation topic. The prompt MUST be in English. Focus on style, high-end design, lighting, and relevant metaphors. NO TEXT in the image." },
        { role: "user", content: `Presentation Topic: "${topic}". Generate a 25-word descriptive prompt.` }
      ],
      temperature: 0.5,
    });
    return res.choices[0].message.content.trim();
  } catch (error) {
    console.error("AI Prompt xatoligi:", error.message);
    return topic + " professional presentation cinematic background illustration";
  }
}

// 🔐 ENV orqali yuklash
const token = process.env.BOT_TOKEN;
const groqApiKey = process.env.GROQ_API_KEY;
const CHANNEL_ID = process.env.CHANNEL_ID || '@salomlarkk';
const ADMIN_ID = parseInt(process.env.ADMIN_ID) || 7561580911;

if (!token || !groqApiKey) {
  console.error("Xatolik: BOT_TOKEN yoki GROQ_API_KEY .env faylida ko'rsatilmagan!");
  process.exit(1);
}

const bot = new TelegramBot(token, { polling: true });
const groq = new Groq({ apiKey: groqApiKey });

// ✅ Obuna tekshiruvi
async function checkSub(userId) {
  try {
    // Handle both @channel and -100xxx ID formats
    const chatMember = await bot.getChatMember(CHANNEL_ID, userId);
    const statuses = ['member', 'administrator', 'creator'];
    return statuses.includes(chatMember.status);
  } catch (error) {
    if (error.response && error.response.statusCode === 400) {
      console.warn(`User ${userId} might not be in the chat correctly.`);
      return false;
    }
    console.error(`Obuna tekshirishda xatolik (${CHANNEL_ID}):`, error.message);
    return false;
  }
}

// 📁 Foydalanuvchilar ma'lumotlarini yuklash
const USERS_FILE = path.join(__dirname, 'users.json');
let userData = {};
if (fs.existsSync(USERS_FILE)) {
  try {
    userData = JSON.parse(fs.readFileSync(USERS_FILE, 'utf8'));
  } catch (e) {
    userData = {};
  }
}

function saveUsers() {
  fs.writeFileSync(USERS_FILE, JSON.stringify(userData, null, 2));
}

// 🎨 Mavzu Ranglari
const THEMES = {
  dark_blue: { bg: '061121', secondary: '00CCFF', accent: '2D68FF', text: 'FFFFFF', subtext: 'EEEEEE' },
  emerald: { bg: '012622', secondary: '00FF99', accent: '00A86B', text: 'FFFFFF', subtext: 'D1FAE5' },
  purple: { bg: '1A0B2E', secondary: 'BF40BF', accent: '702963', text: 'FFFFFF', subtext: 'E6E6FA' },
  gold: { bg: '1A1A1A', secondary: 'D4AF37', accent: 'C5A028', text: 'FFFFFF', subtext: 'F5F5F5' },
  cyberpunk: { bg: '0D0D0D', secondary: 'FF00FF', accent: '00FFFF', text: 'FFFFFF', subtext: 'E0E0E0' },
  ocean: { bg: '003366', secondary: '00CCFF', accent: '007BFF', text: 'FFFFFF', subtext: 'E0FFFF' }
};

// 🤖 Groq orqali slaydlar yaratish (Ko'p tilli qo'llab-quvvatlash)
async function generateSlides(topic) {
  const prompt = `Create exactly 10 extremely detailed slides about "${topic}".
IMPORTANT: Detect the language of the topic and respond ONLY in that language.
Be creative and provide unique content. Do not repeat previous styles.

Use this exact format for each slide:
---
Slide X: Title | Subtitle | Content (at least 3-4 detailed bullet points)
---

Rules:
1. Use '---' to clearly separate each slide block.
2. Use '|' to separate Title, Subtitle, and Content within a slide.
3. In Content, start each point with '-' on a new line. Each point must be 3-4 detailed bullet points.
4. For Slide 1 (Cover), provide a rich description of the topic.
5. Provide COMPLETELY UNIQUE and creative content. Never use the same introduction or structure twice.
6. ONLY return the slide list. Separate each slide block with '---'.`;

  try {
    const res = await groq.chat.completions.create({
      model: "llama-3.3-70b-versatile",
      messages: [
        { role: "system", content: "You are a professional presentation expert. Automatically detect the user's input language and respond in the same language. Use high-quality professional vocabulary. Your content must be rich, unique, and well-formatted." },
        { role: "user", content: prompt }
      ],
      temperature: 0.7,
    });

    const aiText = res.choices[0].message.content;
    console.log("AI Javobi:", aiText);
    return aiText;
  } catch (error) {
    console.error("Groq xatoligi:", error.message);
    throw new Error("AI ma'lumot yarata olmadi.");
  }
}

// 📊 PPTX fayl yasash (High Quality)
async function createPPT(text, filePath, topic, userObj) {
  const pptx = new PptxGenJS();
  pptx.layout = 'LAYOUT_16x9';
  
  const fullName = userObj.name || "Foydalanuvchi";
  const themeKey = userObj.theme || 'dark_blue';
  const theme = THEMES[themeKey] || THEMES.dark_blue;

  // Slaydlarni ajratish (Robust parsing)
  let slidesData = text.split('---').map(s => s.trim()).filter(s => s.includes('|'));

  // Agar '---' bilan ajratishda 1 tadan kam slayd bo'lsa, eski usulda urinib ko'ramiz
  if (slidesData.length <= 1) {
     const fallbackData = text.split(/(?:Slide|Slayd) \d+:\s*/i).map(s => s.trim()).filter(s => s.includes('|'));
     if (fallbackData.length > slidesData.length) {
         slidesData = fallbackData;
     }
  }

  // 1. MUQOVA SLAYDI (Cover)
  const coverSlide = pptx.addSlide();
  coverSlide.background = { color: theme.bg };
  const coverSubtitle = slidesData[0]?.split('|')[1]?.trim() || "Mavzu yuzasidan batafsil ma'lumot";
  const coverImg = userObj.coverImg;

  if (themeKey === 'cyberpunk') {
    // Cyberpunk Decorative (Background)
    coverSlide.addShape(pptx.ShapeType.rect, { x: -0.5, y: 0.5, w: '60%', h: 1.5, fill: { color: theme.secondary }, opacity: 80, rotate: 5 });
    coverSlide.addShape(pptx.ShapeType.rect, { x: '55%', y: 0.3, w: 0.2, h: 2, fill: { color: theme.accent }, rotate: 5 });
    coverSlide.addShape(pptx.ShapeType.rect, { x: '58%', y: 0.2, w: 0.05, h: 2.2, fill: { color: theme.accent }, rotate: 5 });
    coverSlide.addShape(pptx.ShapeType.triangle, { x: 7, y: 4, w: 2, h: 2, fill: { color: theme.accent } });
    coverSlide.addShape(pptx.ShapeType.rect, { x: 0, y: 5.1, w: '100%', h: 0.5, fill: { color: theme.secondary }, opacity: 30 });

    if (coverImg) coverSlide.addImage({ data: coverImg, x: 5.2, y: 1.0, w: 4.2, h: 3.5, sizing: { type: 'cover' } });
    
    coverSlide.addText(topic.toUpperCase(), { x: 0.5, y: 0.7, w: '45%', h: 1, fontSize: 36, bold: true, color: '000000', fontFace: 'Courier New' });
    coverSlide.addText(coverSubtitle, { x: 0.5, y: 2.6, w: '45%', fontSize: 20, color: theme.text, glow: { size: 5, color: theme.accent }, fontFace: 'Courier New' });
    coverSlide.addText(`SYS.ADMIN // ${fullName.toUpperCase()} //`, { x: 0.5, y: 5.2, w: '90%', fontSize: 18, color: theme.secondary, bold: true, fontFace: 'Courier New' });

  } else if (themeKey === 'gold') {
    // Elegant Gold Decorative
    coverSlide.addShape(pptx.ShapeType.rect, { x: '5%', y: '5%', w: '90%', h: '90%', line: { color: theme.secondary, width: 3 } });
    coverSlide.addShape(pptx.ShapeType.rect, { x: '6.5%', y: '8%', w: '87%', h: '84%', line: { color: theme.accent, width: 1 } }); 
    coverSlide.addShape(pptx.ShapeType.diamond, { x: '48.5%', y: '10%', w: 0.3, h: 0.3, fill: { color: theme.secondary } });
    coverSlide.addShape(pptx.ShapeType.diamond, { x: '48.5%', y: '85%', w: 0.3, h: 0.3, fill: { color: theme.secondary } });

    if (coverImg) coverSlide.addImage({ data: coverImg, x: 2.5, y: 1.2, w: 5, h: 2.5, sizing: { type: 'cover' } });

    coverSlide.addText(topic.toUpperCase(), { x: '10%', y: 3.8, w: '80%', h: 0.8, fontSize: 36, bold: true, color: theme.secondary, align: 'center', fontFace: 'Georgia', charSpacing: 3 });
    coverSlide.addText(coverSubtitle, { x: '10%', y: 4.6, w: '80%', fontSize: 18, color: theme.text, align: 'center', italic: true, fontFace: 'Georgia' });
    coverSlide.addText(fullName, { x: '10%', y: 5.1, w: '80%', fontSize: 16, color: theme.accent, align: 'center', charSpacing: 2 });

  } else if (themeKey === 'ocean') {
    // Fluid Ocean Decorative
    coverSlide.addShape(pptx.ShapeType.ellipse, { x: -2, y: 3, w: 12, h: 8, fill: { color: theme.secondary }, opacity: 15 });
    coverSlide.addShape(pptx.ShapeType.ellipse, { x: 3, y: 4, w: 10, h: 7, fill: { color: theme.accent }, opacity: 20 });
    coverSlide.addShape(pptx.ShapeType.ellipse, { x: 8, y: 1, w: 5, h: 5, fill: { color: 'FFFFFF' }, opacity: 10 });

    if (coverImg) coverSlide.addImage({ data: coverImg, x: 0.5, y: 1, w: 4, h: 3.6, sizing: { type: 'cover' } });

    coverSlide.addText(topic, { x: 4.8, y: 1.5, w: '48%', h: 1.5, fontSize: 42, bold: true, color: theme.text, align: 'right', fontFace: 'Trebuchet MS' });
    coverSlide.addText(coverSubtitle, { x: 4.8, y: 3.0, w: '48%', fontSize: 20, color: theme.subtext, align: 'right', fontFace: 'Trebuchet MS' });
    coverSlide.addShape(pptx.ShapeType.rect, { x: '50%', y: 5.1, w: '45%', h: 0.05, fill: { color: theme.secondary }, opacity: 50 });
    coverSlide.addText(fullName, { x: 0.5, y: 5.2, w: '90%', fontSize: 18, color: theme.secondary, align: 'right', fontFace: 'Trebuchet MS' });

  } else if (themeKey === 'emerald') {
    // Emerald layout Decorative
    coverSlide.addShape(pptx.ShapeType.rtTriangle, { x: 4, y: 0, w: 6, h: 5.625, fill: { color: theme.accent }, opacity: 25, flipV: false });
    coverSlide.addShape(pptx.ShapeType.rtTriangle, { x: 4.5, y: 0, w: 5.5, h: 5.625, fill: { color: theme.secondary }, opacity: 15 });
    coverSlide.addShape(pptx.ShapeType.rect, { x: 0.5, y: 5.0, w: 3.5, h: 0.4, fill: { color: theme.secondary }, rectRadius: 0.1 });

    if (coverImg) coverSlide.addImage({ data: coverImg, x: 5.5, y: 1.0, w: 3.5, h: 3.5, sizing: { type: 'cover' } });

    coverSlide.addText(topic.toUpperCase(), { x: 0.5, y: 1.5, w: '45%', h: 2.0, fontSize: 38, bold: true, color: theme.text, fontFace: 'Arial Black' });
    coverSlide.addText(coverSubtitle, { x: 0.5, y: 3.5, w: '45%', fontSize: 18, color: theme.secondary, fontFace: 'Arial' });
    coverSlide.addText(fullName, { x: 0.5, y: 5.0, w: 3.5, h: 0.4, fontSize: 16, color: "000000", align: "center", bold: true });

  } else if (themeKey === 'purple') {
    // Purple layout Decorative
    coverSlide.addShape(pptx.ShapeType.triangle, { x: 5, y: -2, w: 8, h: 10, fill: { color: theme.secondary }, opacity: 40, rotate: 45 });
    coverSlide.addShape(pptx.ShapeType.triangle, { x: 6, y: -1, w: 6, h: 8, fill: { color: theme.accent }, opacity: 50, rotate: 60 });
    coverSlide.addShape(pptx.ShapeType.rect, { x: 9.5, y: 0, w: 0.5, h: '100%', fill: { color: theme.text }, opacity: 10 });
    coverSlide.addShape(pptx.ShapeType.rect, { x: 0.5, y: 5.2, w: 2, h: 0.05, fill: { color: theme.accent } });

    if (coverImg) coverSlide.addImage({ data: coverImg, x: 6.2, y: 1.0, w: 3.2, h: 3.6, sizing: { type: 'cover' } });

    coverSlide.addText(topic.toUpperCase(), { x: 0.5, y: 2.0, w: '55%', h: 1.5, fontSize: 44, bold: true, color: theme.text, align: 'left', shadow: { type: 'outer', color: theme.secondary, blur: 5, offset: 3, angle: 45 } });
    coverSlide.addText(coverSubtitle, { x: 0.5, y: 3.5, w: '55%', fontSize: 20, color: theme.subtext, align: 'left', italic: true });
    coverSlide.addText(`Muallif: ${fullName}`, { x: 0.5, y: 5.3, w: '80%', fontSize: 16, color: theme.accent, align: 'left' });

  } else {
    // Default (Dark Blue Premium) Decorative FIRST (Z-Order)
    coverSlide.addShape(pptx.ShapeType.ellipse, { x: 6.5, y: -1.0, w: 6, h: 6, line: { color: theme.accent, width: 3 }, opacity: 20 });
    coverSlide.addShape(pptx.ShapeType.ellipse, { x: 7.5, y: 1.5, w: 4, h: 4, line: { color: theme.secondary, width: 2 }, opacity: 40 });
    coverSlide.addShape(pptx.ShapeType.ellipse, { x: 8.5, y: 4.5, w: 3, h: 3, fill: { color: theme.accent }, opacity: 15 });
    coverSlide.addShape(pptx.ShapeType.rect, { x: 0.5, y: 4.8, w: 1.5, h: 0.08, fill: { color: theme.accent } });

    coverSlide.addShape(pptx.ShapeType.roundRect, { x: 0.5, y: 0.5, w: 3.5, h: 0.5, fill: { color: theme.secondary }, rectRadius: 0.1 });
    coverSlide.addText(fullName.toUpperCase(), { x: 0.5, y: 0.5, w: 3.5, h: 0.5, fontSize: 16, bold: true, color: '000000', align: 'center' });

    if (coverImg) coverSlide.addImage({ data: coverImg, x: 5.8, y: 1.2, w: 3.8, h: 3.5, sizing: { type: 'cover' } });

    coverSlide.addText(topic.toUpperCase(), { x: 0.5, y: 1.5, w: '50%', h: 1.5, fontSize: 40, bold: true, color: theme.text, align: 'left', fontFace: 'Arial Black', valign: 'top' });
    coverSlide.addText(coverSubtitle, { x: 0.5, y: 3.5, w: '50%', fontSize: 18, color: theme.secondary, italic: true });
  }

  // 2. MA'LUMOTLI SLAYDLAR
  slidesData.forEach((slideBlock, index) => {
    const slide = pptx.addSlide();
    slide.background = { color: theme.bg };

    let [title, subtitle, content] = slideBlock.replace(/(?:Slide|Slayd) \d+:\s*/i, '').split("|").map(s => s?.trim());
    if (!content) {
      content = subtitle;
      subtitle = "";
    }

    const points = content?.split('-').filter(p => p.trim());
    let formattedContent = points?.length > 0 ? points.map(p => p.trim()).join('\n') : (content || "Ma'lumot mavjud emas");

    if (themeKey === 'cyberpunk') {
      slide.addShape(pptx.ShapeType.rect, { x: 0, y: 0.3, w: '65%', h: 1.0, fill: { color: theme.accent }, opacity: 70, rotate: -2 });
      slide.addShape(pptx.ShapeType.rect, { x: '63%', y: 0.2, w: 0.1, h: 1.2, fill: { color: theme.secondary }, rotate: -2 });
      slide.addText(title || "Mavzu", { x: 0.5, y: 0.25, w: '80%', h: 1, fontSize: 32, bold: true, color: '000000', fontFace: 'Courier New' });
      slide.addText(formattedContent, { x: 0.5, y: 1.8, w: '85%', h: 3.2, fontSize: 18, color: theme.text, fontFace: 'Courier New', bullet: { type: 'bullet', characterCode: '25B6' }, lineSpacing: 28 });
      slide.addShape(pptx.ShapeType.rect, { x: 9.5, y: 1.8, w: 0.1, h: 3.2, fill: { color: theme.secondary }, opacity: 50 });
      slide.addText(`PRJ: ${topic.substring(0, 15).toUpperCase()} //`, { x: 0.5, y: 5.2, w: '90%', fontSize: 12, color: theme.accent, fontFace: 'Courier New', opacity: 70 });
    } else if (themeKey === 'gold') {
      slide.addShape(pptx.ShapeType.rect, { x: '4%', y: '4%', w: '92%', h: '92%', line: { color: theme.secondary, width: 1.5 } });
      slide.addShape(pptx.ShapeType.diamond, { x: '49%', y: '4%', w: 0.2, h: 0.2, fill: { color: theme.secondary } });
      slide.addText(title || "Mavzu", { x: 0, y: 0.6, w: '100%', h: 0.8, fontSize: 32, bold: true, color: theme.secondary, align: 'center', fontFace: 'Georgia' });
      slide.addShape(pptx.ShapeType.rect, { x: '40%', y: 1.5, w: '20%', h: 0.02, fill: { color: theme.accent } });
      slide.addText(formattedContent, { x: 1.0, y: 1.8, w: '80%', h: 3.5, fontSize: 18, color: theme.text, fontFace: 'Georgia', align: 'left', lineSpacing: 30, bullet: { type: 'bullet' } });
    } else if (themeKey === 'ocean') {
      slide.addShape(pptx.ShapeType.ellipse, { x: -2, y: -2, w: 5, h: 5, fill: { color: theme.accent }, opacity: 15 });
      slide.addShape(pptx.ShapeType.ellipse, { x: 8, y: 4, w: 6, h: 6, fill: { color: theme.secondary }, opacity: 10 });
      slide.addText(title || "Mavzu", { x: 0.8, y: 0.5, w: '70%', h: 1.0, fontSize: 34, bold: true, color: theme.secondary, fontFace: 'Trebuchet MS' });
      slide.addShape(pptx.ShapeType.rect, { x: 0.8, y: 1.5, w: '15%', h: 0.05, fill: { color: theme.accent } });
      slide.addText(formattedContent, { x: 0.8, y: 1.8, w: '80%', h: 3.5, fontSize: 18, color: theme.text, fontFace: 'Trebuchet MS', lineSpacing: 28, bullet: { type: 'bullet' } });
      slide.addShape(pptx.ShapeType.rect, { x: 0, y: 5.4, w: '100%', h: 0.2, fill: { color: theme.secondary }, opacity: 40 });
    } else if (themeKey === 'emerald') {
      slide.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 1.2, h: '100%', fill: { color: theme.accent }, opacity: 80 });
      slide.addText(String(index + 1).padStart(2, '0'), { x: 0, y: 0.5, w: 1.2, h: 1, fontSize: 24, bold: true, color: '000000', align: 'center', fontFace: 'Arial Black' });
      slide.addText(title || "Mavzu", { x: 1.8, y: 0.5, w: '80%', h: 1, fontSize: 32, bold: true, color: theme.secondary, fontFace: 'Arial' });
      slide.addText(formattedContent, { x: 1.8, y: 1.6, w: '75%', h: 3.8, fontSize: 18, color: theme.text, lineSpacing: 26, bullet: { type: 'number' } });
    } else if (themeKey === 'purple') {
      slide.addShape(pptx.ShapeType.rtTriangle, { x: 8, y: 0, w: 2, h: 3, fill: { color: theme.secondary }, opacity: 20, flipH: true });
      slide.addText(title || "Mavzu", { x: 0.8, y: 0.6, w: '85%', h: 1, fontSize: 36, bold: true, color: theme.text, align: 'left' });
      slide.addShape(pptx.ShapeType.rect, { x: 0.8, y: 1.6, w: 2, h: 0.08, fill: { color: theme.accent } });
      slide.addText(formattedContent, { x: 0.8, y: 2.0, w: '85%', h: 3.5, fontSize: 18, color: theme.subtext, align: 'left', lineSpacing: 28, bullet: { type: 'bullet' } });
    } else {
      // Default (Dark Blue Premium)
      slide.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 0.3, h: '100%', fill: { color: theme.accent } });
      slide.addShape(pptx.ShapeType.rect, { x: 0.3, y: 0, w: 0.1, h: '100%', fill: { color: theme.secondary }, opacity: 50 });
      slide.addText(title || "Mavzu", { x: 0.8, y: 0.4, w: '85%', h: 0.8, fontSize: 30, bold: true, color: theme.secondary, fontFace: 'Arial Black' });
      if (subtitle) {
        slide.addText(subtitle, { x: 0.8, y: 1.2, w: '80%', fontSize: 16, color: theme.accent, italic: true });
      }
      slide.addShape(pptx.ShapeType.rect, { x: 0.8, y: 1.7, w: 3, h: 0.05, fill: { color: theme.secondary } });
      slide.addText(formattedContent, { x: 0.8, y: 2.0, w: '85%', h: 3.3, fontSize: 18, color: theme.text, bullet: { type: 'bullet', characterCode: '2192' }, valign: 'top', lineSpacing: 28 });
      slide.addShape(pptx.ShapeType.ellipse, { x: 9.5, y: 5.0, w: 2, h: 2, fill: { color: theme.accent }, opacity: 10 });
      slide.addShape(pptx.ShapeType.triangle, { x: 10.5, y: -0.5, w: 2, h: 2, fill: { color: theme.secondary }, opacity: 15, rotate: 180 });
      slide.addText(`@salomlarkk`, { x: 0.5, y: 5.4, w: '90%', fontSize: 11, color: theme.subtext, opacity: 50, align: 'right' });
    }
  });

  await pptx.writeFile({ fileName: filePath });
  return filePath;
}

// Foydalanuvchi holatlari
const userStates = {};

// 📊 Admin Buyruqlari: Statistika
bot.onText(/\/stat/, (msg) => {
  if (msg.from.id !== ADMIN_ID) return;
  const count = Object.keys(userData).length;
  bot.sendMessage(msg.chat.id, `📊 <b>Bot Statistikasi:</b>\n\nJami foydalanuvchilar: <b>${count}</b> ta`, { parse_mode: 'HTML' });
});

// 📊 Admin Buyruqlari: Foydalanuvchilar Ro'yxati
bot.onText(/\/users/, (msg) => {
  if (msg.from.id !== ADMIN_ID) return;
  let list = "📋 <b>Foydalaruvchilar Ro'yxati:</b>\n\n";
  const ids = Object.keys(userData);
  if (ids.length === 0) list += "<i>Hali foydalanuvchilar yo'q.</i>";
  ids.forEach((id, index) => {
    const user = userData[id];
    const name = typeof user === 'object' ? user.name : user;
    list += `${index + 1}. <a href="tg://user?id=${id}">${id}</a> - <b>${name}</b>\n`;
  });
  bot.sendMessage(msg.chat.id, list, { parse_mode: 'HTML' });
});

// 📢 Admin Buyruqlari: Xabar yuborish (Broadcast)
bot.onText(/\/send (.+)/, async (msg, match) => {
  if (msg.from.id !== ADMIN_ID) return;
  const text = match[1];
  const ids = Object.keys(userData);
  
  bot.sendMessage(msg.chat.id, `🚀 <b>${ids.length}</b> ta foydalanuvchiga xabar yuborish boshlandi...`, { parse_mode: 'HTML' });

  let success = 0;
  for (const id of ids) {
    try {
      await bot.sendMessage(id, text, { parse_mode: 'HTML' });
      success++;
    } catch (e) {
      console.error(`Xabar yuborishda xatolik (${id}):`, e.message);
    }
    await new Promise(res => setTimeout(res, 50)); // Rate limit protection
  }

  bot.sendMessage(msg.chat.id, `✅ Xabar yuborish tugadi. \n\nYetib bordi: <b>${success}</b> ta.`, { parse_mode: 'HTML' });
});

// 🔄 Ismni o'zgartirish
bot.onText(/\/resetname/, (msg) => {
  const userId = msg.from.id;
  bot.sendMessage(userId, "👤 Yangi ism va familiyangizni yuboring:");
  userStates[userId] = 'waiting_name';
});

// 🎨 Mavzu tanlash
bot.onText(/\/theme/, (msg) => {
  const userId = msg.from.id;
  bot.sendMessage(userId, "🎨 Slayd dizayni uchun mavzu tanlang:", {
    reply_markup: {
      inline_keyboard: [
        [{ text: "🌑 To'q Ko'k (Premium)", callback_data: "theme_dark_blue" }],
        [{ text: "🌿 Zumrad Yashil", callback_data: "theme_emerald" }],
        [{ text: "🔮 Binafsha (Premium)", callback_data: "theme_purple" }],
        [{ text: "✨ Oltin (Elegance)", callback_data: "theme_gold" }],
        [{ text: "⚡ Cyberpunk (Futuristic)", callback_data: "theme_cyberpunk" }],
        [{ text: "🌊 Okean (Tinchlik)", callback_data: "theme_ocean" }]
      ]
    }
  });
});

// 🚀 START komandasi
bot.onText(/\/start/, async (msg) => {
  const userId = msg.from.id;
  const ok = await checkSub(userId);

  if (!ok) {
    // 🔔 Adminga xabar yuborish (Obuna bo'lmagan foydalanuvchi)
    const userLink = msg.from.username ? `<a href="https://t.me/${msg.from.username}">@${msg.from.username}</a>` : `<a href="tg://user?id=${userId}">ID: ${userId}</a>`;
    bot.sendMessage(ADMIN_ID, `⚠️ <b>Potensial Foydalanuvchi (Obunasiz):</b> \n👤 Ism: <b>${msg.from.first_name} ${msg.from.last_name || ''}</b>\n🔗 Profil: ${userLink}\n🆔 ID: <code>${userId}</code>`, { parse_mode: 'HTML' }).catch(e => console.error("Admin notification error:", e.message));

    return bot.sendMessage(userId, `❌ Botdan foydalanish uchun <b>${CHANNEL_ID}</b> kanaliga obuna bo'ling!`, {
      parse_mode: 'HTML',
      reply_markup: {
        inline_keyboard: [
          [{ text: "📢 Kanalga o'tish", url: `https://t.me/${CHANNEL_ID.startsWith('@') ? CHANNEL_ID.replace('@','') : 'salomlarkk'}` }],
          [{ text: "✅ Tekshirish", callback_data: "check_sub" }]
        ]
      }
    });
  }

  const user = userData[userId];
  if (user && (typeof user === 'object' ? user.name : user)) {
    const name = typeof user === 'object' ? user.name : user;
    return bot.sendMessage(userId, `👋 Assalomu alaykum, <b>${name}</b>! \n\nSiz ro'yxatdan o'tgansiz.\n\n🎨 Mavzuni o'zgartirish uchun: /theme\n👤 Ismni o'zgartirish uchun: /resetname\n\nSlayd yaratish uchun mavzu yuboring:`, { parse_mode: 'HTML' });
  }

  bot.sendMessage(userId, "👤 Slaydda chiqarish uchun ism va familiyangizni yuboring:");
  userStates[userId] = 'waiting_name';

  // 🔔 Adminga xabar yuborish (Yangi boshlovchi)
  const userLink = msg.from.username ? `<a href="https://t.me/${msg.from.username}">@${msg.from.username}</a>` : `<a href="tg://user?id=${userId}">ID: ${userId}</a>`;
  bot.sendMessage(ADMIN_ID, `🆕 <b>Yangi Foydalanuvchi (Start bosdi):</b> \n👤 Ism: <b>${msg.from.first_name} ${msg.from.last_name || ''}</b>\n🔗 Profil: ${userLink}\n🆔 ID: <code>${userId}</code>`, { parse_mode: 'HTML' }).catch(e => console.error("Admin notification error:", e.message));
});

// 🔁 Tekshirish va Mavzu tugmalari
bot.on('callback_query', async (q) => {
  const userId = q.from.id;
  
  if (q.data === "check_sub") {
    const ok = await checkSub(userId);
    if (ok) {
      const user = userData[userId];
      if (user && (typeof user === 'object' ? user.name : user)) {
        bot.sendMessage(userId, "✅ Obuna tasdiqlandi. Slayd mavzusini yuboring:");
      } else {
        bot.sendMessage(userId, "✅ Endi ism va familiyangizni yuboring:");
        userStates[userId] = 'waiting_name';
      }
    } else {
      bot.answerCallbackQuery(q.id, {
        text: "❌ Hali obuna bo'lmagansiz!",
        show_alert: true
      });
    }
  }

  if (q.data.startsWith('theme_')) {
    const themeKey = q.data.replace('theme_', '');
    if (!userData[userId]) userData[userId] = { name: "", theme: "dark_blue" };
    if (typeof userData[userId] === 'string') userData[userId] = { name: userData[userId], theme: "dark_blue" };
    
    userData[userId].theme = themeKey;
    saveUsers();
    
    bot.answerCallbackQuery(q.id, { text: "🎨 Mavzu saqlandi!" });
    bot.sendMessage(userId, `✅ Slayd dizayni <b>${themeKey}</b> mavzusiga o'zgartirildi!`, { parse_mode: 'HTML' });
  }
});

// 📩 Xabarlarni qayta ishlash
bot.on('message', async (msg) => {
  if (!msg.text || msg.text.startsWith('/')) return;

  const chatId = msg.chat.id;
  const userId = msg.from.id;

  // Ism yozishni kutish
  if (userStates[userId] === 'waiting_name') {
    const fullName = msg.text.trim();
    if (!userData[userId]) userData[userId] = { name: fullName, theme: "dark_blue" };
    else if (typeof userData[userId] === 'object') userData[userId].name = fullName;
    else userData[userId] = { name: fullName, theme: "dark_blue" };
    
    saveUsers();
    delete userStates[userId];

    // 🔔 Adminga xabar yuborish
    const userLink = msg.from.username ? `<a href="https://t.me/${msg.from.username}">@${msg.from.username}</a>` : `<a href="tg://user?id=${userId}">ID: ${userId}</a>`;
    bot.sendMessage(ADMIN_ID, `✅ <b>Ism Saqlandi:</b> \n👤 Ism: <b>${fullName}</b>\n🔗 Profil: ${userLink}\n🆔 ID: <code>${userId}</code>`, { parse_mode: 'HTML' }).catch(e => console.error("Admin notification error:", e.message));

    return bot.sendMessage(chatId, `✅ Ism saqlandi: <b>${fullName}</b>\n\nEndi slayd mavzusini yuboring:`, { parse_mode: 'HTML' });
  }

  const ok = await checkSub(userId);
  if (!ok) return;

  if (!userData[userId]) {
    bot.sendMessage(chatId, "👤 Iltimos, darslik uchun ism va familiyangizni yozib yuboring:");
    userStates[userId] = 'waiting_name';
    return;
  }

  const statusMsg = await bot.sendMessage(chatId, "⏳ Ma'lumotlar to'planmoqda va premium dizayn yaratilmoqda...");

  try {
    const imgPrompt = await generateImagePrompt(msg.text);
    const aiText = await generateSlides(msg.text);
    const coverImg = await fetchImageBase64(imgPrompt);

    const filePath = path.join(__dirname, `slayd_${chatId}_${Date.now()}.pptx`);

    const userObj = typeof userData[userId] === 'object' ? userData[userId] : { name: userData[userId], theme: "dark_blue" };
    await createPPT(aiText, filePath, msg.text, { ...userObj, coverImg });

    await bot.sendDocument(chatId, filePath, { caption: `✅ \"<b>${msg.text}</b>\" mavzusidagi mukammal slayd tayyor!\n\n🎨 Mavzu: <b>${userObj.theme || 'dark_blue'}</b>\n👤 Tayyorladi: <b>${userObj.name}</b>\n\n📢 Kanal: ${CHANNEL_ID}`, parse_mode: 'HTML' });

    if (fs.existsSync(filePath)) {
      fs.unlinkSync(filePath);
    }
    
    bot.deleteMessage(chatId, statusMsg.message_id).catch(() => {});

  } catch (error) {
    console.error("Xatolik:", error);
    bot.sendMessage(chatId, "❌ Kechirasiz, xatolik yuz berdi: " + error.message);
  }
});

// 🌐 Render uchun oddiy veb-server (Health Check)
const PORT = process.env.PORT || 3000;
app.get('/', (req, res) => {
  res.send('Bot is running...');
});

app.listen(PORT, () => {
  console.log(`Web server is listening on port ${PORT}`);
});
