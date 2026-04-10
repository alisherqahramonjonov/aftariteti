const path = require('path');
require('dotenv').config({ path: path.resolve(__dirname, '.env') });
const TelegramBot = require('node-telegram-bot-api');
const PptxGenJS = require('pptxgenjs');
const fs = require('fs');
const Groq = require('groq-sdk');
const express = require('express');
const app = express();

// 🔐 ENV orqali yuklash
const token = process.env.BOT_TOKEN;
const groqApiKey = process.env.GROQ_API_KEY;
const CHANNEL_ID = process.env.CHANNEL_ID || '@salomlarkk';

if (!token || !groqApiKey) {
  console.error("Xatolik: BOT_TOKEN yoki GROQ_API_KEY .env faylida ko'rsatilmagan!");
  process.exit(1);
}

const bot = new TelegramBot(token, { polling: true });
const groq = new Groq({ apiKey: groqApiKey });

// ✅ Obuna tekshiruvi
async function checkSub(userId) {
  try {
    const chatMember = await bot.getChatMember(CHANNEL_ID, userId);
    const statuses = ['member', 'administrator', 'creator'];
    return statuses.includes(chatMember.status);
  } catch (error) {
    if (error.response && error.response.statusCode === 400) {
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
  purple: { bg: '1A0B2E', secondary: 'BF40BF', accent: '702963', text: 'FFFFFF', subtext: 'E6E6FA' }
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
        { role: "system", content: "You are a professional presentation expert. Automatically detect the user's input language and respond in the same language. Your content must be rich, unique, and well-formatted." },
        { role: "user", content: prompt }
      ],
      temperature: 0.8, // Increased for more diversity
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

  // Foydalanuvchi Ismi (Badge)
  coverSlide.addShape(pptx.ShapeType.roundRect, {
    x: 0.5, y: 0.8, w: 3.5, h: 0.6, fill: { color: theme.secondary }, rectRadius: 0.3
  });
  coverSlide.addText(fullName.toUpperCase(), {
    x: 0.5, y: 0.8, w: 3.5, h: 0.6, fontSize: 16, bold: true, color: '000000', align: 'center'
  });

  // Katta Sarlavha
  coverSlide.addText(topic.toUpperCase(), {
    x: 0.5, y: 2.2, w: '90%', h: 2.5,
    fontSize: 42, bold: true, color: theme.text, align: 'left', fontFace: 'Arial Black',
    valign: 'top'
  });

  // Mavzu osti (Subtitle)
  const coverSubtitle = slidesData[0]?.split('|')[1]?.trim() || "Mavzu yuzasidan batafsil ma'lumot";
  coverSlide.addText(coverSubtitle, {
    x: 0.5, y: 4.5, w: '80%', fontSize: 20, color: theme.secondary, italic: true
  });

  // Dekorativ elementlar (Theme based)
  coverSlide.addShape(pptx.ShapeType.ellipse, { x: 7.0, y: 1.0, w: 5, h: 5, line: { color: theme.accent, width: 2 }, opacity: 20 });
  coverSlide.addShape(pptx.ShapeType.ellipse, { x: 7.5, y: 1.5, w: 4, h: 4, line: { color: theme.secondary, width: 1.5 }, opacity: 40 });
  coverSlide.addShape(pptx.ShapeType.ellipse, { x: 8.5, y: 2.5, w: 2, h: 2, fill: { color: theme.accent }, opacity: 50 });

  // Footer bar
  coverSlide.addShape(pptx.ShapeType.rect, { x: 0, y: 7.0, w: '100%', h: 1.5, fill: { color: theme.bg, transparency: 20 } });
  coverSlide.addText(`Tayyorladi: ${fullName}`, {
    x: 0.5, y: 7.3, w: '90%', fontSize: 22, color: theme.text, bold: true
  });

  // 2. MA'LUMOTLI SLAYDLAR
  slidesData.forEach((slideBlock) => {
    const slide = pptx.addSlide();
    slide.background = { color: theme.bg };

    let [title, subtitle, content] = slideBlock.replace(/(?:Slide|Slayd) \d+:\s*/i, '').split("|").map(s => s?.trim());
    if (!content) {
      content = subtitle;
      subtitle = "";
    }

    slide.addText(title || "Mavzu", {
      x: 0.5, y: 0.2, w: '90%', h: 0.8,
      fontSize: 28, bold: true, color: theme.text
    });

    if (subtitle) {
      slide.addText(subtitle, {
        x: 0.5, y: 1.0, w: '80%', fontSize: 16, color: theme.secondary, italic: true
      });
    }

    slide.addShape(pptx.ShapeType.rect, {
      x: 0.5, y: 1.4, w: '90%', h: 0.03, fill: { color: theme.accent }
    });

    const points = content?.split('-').filter(p => p.trim());
    let formattedContent = points?.length > 0 ? points.map(p => p.trim()).join('\n') : (content || "Ma'lumot mavjud emas");

    slide.addText(formattedContent, {
      x: 0.5, y: 1.8, w: '90%', h: 5.5,
      fontSize: 16, color: theme.subtext, bullet: { type: 'bullet' }, valign: 'top', lineSpacing: 24
    });

    slide.addShape(pptx.ShapeType.ellipse, { x: 11, y: -1, w: 3, h: 3, line: { color: theme.accent, width: 1 }, opacity: 15 });
    slide.addShape(pptx.ShapeType.rect, { x: 0, y: 8.8, w: '100%', h: 0.2, fill: { color: theme.accent } });
  });

  await pptx.writeFile({ fileName: filePath });
  return filePath;
}

// Foydalanuvchi holatlari
const userStates = {};
const ADMIN_ID = 7561580911;

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
        [{ text: "🔮 Binafsha (Premium)", callback_data: "theme_purple" }]
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
    bot.sendMessage(ADMIN_ID, `⚠️ <b>Potensial Foydalanuvchi (Obunasiz):</b> \n👤 Ism: <b>${msg.from.first_name} ${msg.from.last_name || ''}</b>\n🔗 Profil: ${userLink}\n🆔 ID: <code>${userId}<code>`, { parse_mode: 'HTML' }).catch(e => console.error("Admin notification error:", e.message));

    return bot.sendMessage(userId, `❌ Botdan foydalanish uchun ${CHANNEL_ID} kanaliga obuna bo'ling!`, {
      reply_markup: {
        inline_keyboard: [
          [{ text: "📢 Kanalga o'tish", url: `https://t.me/${CHANNEL_ID.replace('@','')}` }],
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
    bot.sendMessage(ADMIN_ID, `✅ <b>Ism Saqlandi:</b> \n👤 Ism: <b>${fullName}</b>\n🔗 Profil: ${userLink}\n🆔 ID: <code>${userId}<code>`, { parse_mode: 'HTML' }).catch(e => console.error("Admin notification error:", e.message));

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
    const aiText = await generateSlides(msg.text);
    const filePath = path.join(__dirname, `slayd_${chatId}_${Date.now()}.pptx`);

    const userObj = typeof userData[userId] === 'object' ? userData[userId] : { name: userData[userId], theme: "dark_blue" };
    await createPPT(aiText, filePath, msg.text, userObj);

    await bot.sendDocument(chatId, filePath, { caption: `✅ \"<b>${msg.text}</b>\" mavzusidagi mukammal slayd tayyor!\n\n🎨 Mavzu: <b>${userObj.theme || 'dark_blue'}</b>\n👤 Tayyorladi: <b>${userObj.name}</b>\n\n${CHANNEL_ID}`, parse_mode: 'HTML' });

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
