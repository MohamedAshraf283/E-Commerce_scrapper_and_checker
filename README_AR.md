# Salla / Zid Desktop Scraper

برنامج مكتبي Python/Tkinter لفحص وسحب بيانات متاجر Salla و Zid من ملف Excel.

## ما الذي يفعله؟

- يقرأ ملف Excel ويكتشف عمود روابط المتاجر أو يسمح باختياره يدويًا.
- يستخدم Playwright لتشغيل متصفح Chromium/Brave بجلسة Persistent Context.
- يفحص حالة المتجر: يعمل، لا يعمل، غير مؤكد، تحقق بشري، خطأ DNS/Timeout/SSL.
- يستخرج بيانات التواصل من الهيدر والفوتر: الإيميلات، الجوالات، واتساب، السوشيال، العنوان إن وُجد.
- يسحب المنتجات من التصنيفات المكتشفة تلقائيًا أو من روابط تصنيفات موجودة في الشيت.
- يدعم Selectors مخصصة للمنتجات والتصنيفات.
- يحفظ النتائج في Excel يحتوي:
  - `Stores`
  - `Products`
  - `Summary`
- يحتوي أداة مستقلة لترتيب ملف Excel بناءً على عمود محدد أبجديًا أو رقميًا.

## التشغيل السريع على Windows

1. فك الضغط عن المشروع.
2. شغل:

```bat
install_and_run.bat
```

سيقوم بإنشاء بيئة افتراضية وتثبيت المتطلبات وتشغيل البرنامج.

بعد أول تثبيت يمكنك تشغيل:

```bat
run_app.bat
```

## إعداد جلسة Brave / Chrome المحلية

في الواجهة:

- `مسار المتصفح exe`: مثال:
  `C:\Program Files\BraveSoftware\Brave-Browser\Application\brave.exe`
- `مجلد User Data`: مثال:
  `C:\Users\YOUR_USER\AppData\Local\BraveSoftware\Brave-Browser\User Data`
- `اسم Profile`: غالبًا `Default` أو `Profile 1`

مهم: لا تضع `User Data\Default` في خانة User Data. ضع مجلد `User Data` فقط، ثم اكتب `Default` في خانة Profile.

يفضل إغلاق Brave/Chrome قبل تشغيل نفس البروفايل داخل البرنامج، لأن بعض المتصفحات تقفل ملفات البروفايل أثناء التشغيل.

## شكل ملف Excel المصدر

أهم عمود مطلوب هو عمود روابط المتاجر. يمكن أن يكون اسمه مثل:

- `link`
- `url`
- `رابط`
- `رابط المتجر`

ويمكن اختيار العمود يدويًا من تبويب `عرض الشيت`.

اختياريًا يمكن إضافة عمود يحتوي روابط التصنيفات، وفي كل خلية يمكن وضع أكثر من رابط مفصولًا بسطر جديد أو `;` أو `|`.

## Selectors JSON

اترك القيم فارغة للاكتشاف التلقائي:

```json
{
  "category_links": "",
  "product_card": "",
  "product_name": "",
  "product_price": "",
  "product_url": "",
  "product_image": ""
}
```

مثال مخصص:

```json
{
  "category_links": "nav a, .categories a",
  "product_card": ".product-card",
  "product_name": ".product-title",
  "product_price": ".price",
  "product_url": "a",
  "product_image": "img"
}
```

## التحقق البشري

البرنامج لا يتجاوز CAPTCHA ولا يحاول حلها. عند اكتشاف صفحة تحقق بشري، يغلق الجلسة وينتظر المدة المحددة في الواجهة ثم يفتح جلسة جديدة ويعيد المحاولة حسب عدد المحاولات المحدد.

## بناء نسخة exe

شغل:

```bat
build_exe.bat
```

النتيجة ستكون داخل:

```text
dist\SallaZidScraper
```

ملاحظة: Playwright يحتاج تثبيت Chromium مرة واحدة على الجهاز:

```bat
python -m playwright install chromium
```

## ملاحظات مهمة

- السحب من المتاجر يجب أن يتم بما يتوافق مع شروط المواقع والقوانين المحلية.
- ابدأ بتوازي منخفض `1` أو `2`.
- لا توجد طريقة عامة تضمن استخراج المنتجات من كل قالب متجر، لذلك وفرت الواجهة Selectors مخصصة لكل حالة.
- عند استخدام جلسة محلية حقيقية، لا تستخدم نفس البروفايل المفتوح في متصفح آخر في نفس اللحظة.
