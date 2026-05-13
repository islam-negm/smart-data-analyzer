# 📊 Smart Data Analyzer

نظام ذكي متكامل لتحليل البيانات تلقائياً وتحويلها إلى تقارير احترافية وقرارات عملية.

---

## ✨ المميزات

- 📥 **جلب تلقائي** لملفات Excel من Google Drive
- 🧹 **تنظيف ذكي** للبيانات (إزالة التكرار، ملء القيم المفقودة)
- 📊 **تحليل شامل**: مبيعات، مالية، تسويق، موارد بشرية
- 📈 **رسوم بيانية متعددة**: Bar, Line, Pie, Scatter, Heatmap, Dashboard
- 🤖 **تقرير عربي تلقائي** بأهم المؤشرات والتوصيات
- 📄 **PDF احترافي** بالتقرير الكامل والرسوم
- 📊 **Excel بالنتائج** مع إحصاء تفصيلي لكل ورقة
- 🎨 **PowerPoint** عرض تقديمي جاهز
- 🌐 **Streamlit Dashboard** تفاعلي
- 📧 **إرسال تلقائي** للتقارير بالبريد الإلكتروني

---

## 🚀 التثبيت

```bash
git clone https://github.com/username/smart-data-analyzer.git
cd smart-data-analyzer
pip install -r requirements.txt
```

---

## 📖 طريقة الاستخدام

### تحليل ملف Excel محلي
```bash
python smart_data_analyzer.py ملفك.xlsx
```

### تشغيل Dashboard التفاعلي
```bash
streamlit run streamlit_dashboard.py
```

### مراقبة Google Drive تلقائياً
```bash
# 1. ضع ملف credentials.json في نفس المجلد
# 2. عدّل WATCHED_FOLDER في google_drive_watcher.py
python google_drive_watcher.py
```

### إرسال التقارير بالإيميل
```bash
export EMAIL_SENDER="your@email.com"
export EMAIL_PASSWORD="app_password"
export EMAIL_RECIPIENT="recipient@email.com"
python email_notifier.py
```

---

## 📦 المخرجات

بعد كل تحليل يتم إنشاء مجلد `outputs/` يحتوي على:

| الملف | الوصف |
|---|---|
| `report_TIMESTAMP.pdf` | تقرير PDF احترافي |
| `results_TIMESTAMP.xlsx` | Excel بالنتائج الكاملة |
| `presentation_TIMESTAMP.pptx` | عرض PowerPoint |
| `*_TIMESTAMP.png` | الرسوم البيانية |
| `report_TIMESTAMP.txt` | التقرير النصي بالعربية |

---

## 🏗️ هيكل المشروع

```
smart-data-analyzer/
│
├── smart_data_analyzer.py     ← النظام الرئيسي (11 وحدة)
├── streamlit_dashboard.py     ← Dashboard التفاعلي
├── google_drive_watcher.py    ← مراقب Google Drive
├── email_notifier.py          ← مُرسل التقارير
├── requirements.txt           ← المكتبات المطلوبة
├── .gitignore                 ← الملفات المستثناة
└── README.md                  ← هذا الملف
```

---

## ⚙️ إعداد Google Drive

1. اذهب إلى [Google Cloud Console](https://console.cloud.google.com)
2. أنشئ مشروعاً جديداً وفعّل **Google Drive API**
3. أنشئ **OAuth 2.0 credentials** وحمّل `credentials.json`
4. ضع الملف في نفس مجلد المشروع
5. عدّل `WATCHED_FOLDER` بمعرّف المجلد المطلوب مراقبته

> ⚠️ **لا ترفع `credentials.json` أو `token.pickle` على GitHub أبداً**

---

## 🧰 المكتبات المستخدمة

| المكتبة | الاستخدام |
|---|---|
| `pandas` | تحليل ومعالجة البيانات |
| `numpy` | العمليات الرياضية |
| `matplotlib` | الرسوم البيانية |
| `scikit-learn` | التحليل الإحصائي والتجميع |
| `reportlab` | إنشاء تقارير PDF |
| `python-pptx` | عروض PowerPoint |
| `openpyxl` | ملفات Excel |
| `streamlit` | Dashboard التفاعلي |

---

## 📋 المتطلبات

- Python 3.9+
- اتصال بالإنترنت (لـ Google Drive فقط)

---

## 📄 الرخصة

MIT License – حر الاستخدام والتعديل

---

> صُنع بـ ❤️ | Smart Data Analyzer © 2025
