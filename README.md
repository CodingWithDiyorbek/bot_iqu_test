# IQ Test Telegram Bot

Python + Telegram Bot API + PostgreSQL (Railway). Lokal sinov uchun SQLite ham ishlaydi. Ro‘yxatdan o‘tish (maktab va sinf bilan), admin belgilagan vaqtda boshlanadigan IQ test, natija, PDF sertifikat (Al-Aziz Academyasi nomidan), admin uchun Excel va sertifikatlar ZIP eksporti.

## Railway'ga joylash (24/7 ishlashi uchun)

1. **GitHub'ga yuklang.** github.com da yangi **Private** repository oching va ZIP ichidagi barcha fayllarni yuklang. `.env` faylini yuklamang.
2. **Railway'da loyiha oching.** railway.com → **New Project** → **Deploy from GitHub repo** → repositoryingizni tanlang. Railway `Dockerfile` va `railway.json` ni o‘zi topadi.
3. **PostgreSQL qo‘shing.** Loyiha ichida **+ New** → **Database** → **Add PostgreSQL**.
4. **O‘zgaruvchilarni kiriting.** Bot servisini bosing → **Variables** → quyidagilarni qo‘shing:

```
BOT_TOKEN=BotFather_bergan_token
ADMIN_LOGIN=admin
ADMIN_PASSWORD=kuchli_parol
DATABASE_URL=${{Postgres.DATABASE_URL}}
```

   `DATABASE_URL` qiymatini aynan shunday yozing, Railway uni PostgreSQL manziliga o‘zi almashtiradi.
5. **Deploy.** O‘zgaruvchilar saqlangach bot qayta ishga tushadi. **Deployments → View logs** da `Baza: PostgreSQL` va `IQ test bot ishga tushdi` yozuvlari chiqsa, hammasi joyida.

Muhim: bot Railway'da ishlayotganda kompyuteringizdagi botni o‘chiring. Bitta token bilan ikkita bot ishlasa Telegram `Conflict` xatosini beradi.
Kodni o‘zgartirsangiz, GitHub'ga yuklashingiz bilan Railway botni avtomatik yangilaydi. Ma’lumotlar PostgreSQL'da saqlanadi va yangilanishda o‘chib ketmaydi.

## Windows / VS Code orqali ishga tushirish

1. ZIP faylni papkaga chiqaring, masalan `C:\Users\User\Desktop\icu_test_telegram_bot`.
2. Python 3.11 yoki undan yangiroq versiyasini o‘rnating. O‘rnatishda **Add Python to PATH** ni belgilang.
3. Papkani VS Code'da oching. Terminal → New Terminal ni bosing.
4. Quyidagilarni birma-bir kiriting:

```powershell
py -m venv venv
.\venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
pip install -r requirements.txt
Copy-Item .env.example .env
notepad .env
```

5. `.env` ichida `BOT_TOKEN=` qatoridan keyin BotFather bergan tokenni qo‘ying. `ADMIN_LOGIN` va `ADMIN_PASSWORD` ni ham o‘zingiz belgilang. Parolni boshqalarga bermang. Faylni saqlang.
6. Botni ishga tushiring:

```powershell
python main.py
```

Yoki Windows'da `run_windows.bat` faylini ikki marta bosing.

## BotFather tokenini qayerga qo‘yaman?
`.env` faylidagi quyidagi qatorga:

`BOT_TOKEN=PASTE_YOUR_TELEGRAM_BOT_TOKEN_HERE`

Tenglik belgisidan keyin tokenni yozing, qo‘shtirnoq shart emas. Tokenni chatga yubormang va GitHub'ga joylamang.

## Foydalanuvchi oqimi
- /start → “Ro‘yxatdan o‘tish” yoki “Admin panelga kirish”.
- Ism-familya → telefon (Telegram kontakt tugmasi orqali) → maktab nomi (yoziladi) → sinf (1–11).
- “IQ testni boshlash” tugmasi admin belgilagan kun va vaqtdan oldin bosilsa, bot “IQ test 20.10.2026 kuni soat 10:00 da boshlanadi” deb javob beradi.
- Belgilangan vaqt kelganda bot ro‘yxatdan o‘tgan hammaga “IQ test boshlandi!” xabarini yuboradi.
- Har bir test 30 savol, 3 variantdan biri tasodifan beriladi. Har bir ishtirokchi testni bir marta topshiradi.
- Test paytida pastda “🏁 Testni yakunlash” tugmasi turadi. Bosilganda tasdiq so‘raladi, javob berilmagan savollar noto‘g‘ri hisoblanadi.
- Test tugagach ball, foiz, izoh va ism-familya, foiz va izoh yozilgan PDF sertifikat yuboriladi.

## Admin panel
`.env` dagi login va parol bilan kiring (parol yozilgan xabar chatdan o‘chiriladi). Bot qayta ishga tushsa ham admin kirgan holatda qoladi.
- ⏰ Test vaqtini belgilash: masalan `20.10.2026 10:00` (Toshkent vaqti). Belgilangach ro‘yxatdan o‘tganlarga vaqt haqida xabar boradi.
- 🗑 Test vaqtini bekor qilish.
- 📊 Statistika, 👥 Foydalanuvchilar (hammasi, ism, telefon, maktab, sinf, natija).
- 📥 Excel: barcha ma’lumotlar, foiz va izoh bilan.
- 🏅 Sertifikatlar: testni tugatganlarning barcha sertifikatlari bitta ZIP faylda.

## Muhim
- Lokal ishga tushirilganda bot faqat kompyuter yoqiq paytda ishlaydi. Test vaqtida bot albatta ishlab turishi kerak, shuning uchun Railway tavsiya etiladi.
- Sertifikat shriftlari `fonts/` papkasida (Montserrat, bepul OFL litsenziya), shuning uchun har qanday kompyuterda bir xil ko‘rinadi.
- Savollar main.py ichidagi QUESTION_BANK da. Ishlatishdan oldin o‘qituvchi tekshirib chiqishi tavsiya etiladi.
