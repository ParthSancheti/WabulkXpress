python -m PyInstaller --onefile --windowed --icon=bin/loco.ico WaBulkSender.py
Absolutely! Here’s a **ready-to-paste, pro-level README.md** for your WhatsApp Bulk Automation repo.

---

# 🚀 WabulkXpress - WhatsApp Bulk Message/Attachment Automation

A **free, full-featured, and blazing-fast WhatsApp bulk messaging/attachment sender** for Windows/Linux, with Excel/CSV import, CLI, robust retry logic, HTML analytics, auto-number formatting, and ZERO pre-setup required.
**No paid API, no official API, no monthly fees.**

---

## ⚡ Features

* **One-file, pure Python:** Easy to audit, edit, or package as EXE.
* **No browser setup needed:** Uses [webdriver-manager](https://github.com/SergeyPirogov/webdriver_manager) to auto-install Chrome driver if needed.
* **Bulk send via Excel or CSV:** Reads numbers from any column, skips header automatically.
* **Send text, attachments, or both:** Supports PDFs, images, videos, audio, docs, etc.
* **Retry logic:** Each send retries up to 3 times per number, with random delay (anti-ban).
* **All messages/attachments are logged with status and analytics.**
* **Automatic phone number normalization:** Handles +91, +1, +44, etc., and pads missing codes.
* **Full HTML report:** See success/failure breakdown after any run.
* **Emoji-rich console feedback.**
* **Works on Windows/Linux/Mac (needs Chrome installed).**
* **NO WhatsApp Business API needed, no cost.**

---

## 📦 Requirements

```sh
pip install selenium webdriver-manager openpyxl pandas
```

* **Google Chrome** installed (any stable version).

---

## 🔥 How To Use

**Login (first time, or to reset session):**

```sh
python wa.py login
```

### Bulk Messages (no attachment):

```sh
python wa.py exl -exl contacts.xlsx -col A "Hi from Python bot!"
python wa.py exl -exl contacts.csv -col mobile "Hello CSV friends!"
```

### Bulk Messages **with** Attachment(s):

```sh
python wa.py exl_atg -exl contacts.xlsx -col A -fileloc "C:\file.pdf" "See attached doc"
python wa.py exl_atg -exl contacts.csv -col mobile -fileloc "C:\img1.jpg,C:\img2.jpg" "Pics!"
```

* **Multiple files** = comma-separated list.
* Will use the first file for all numbers (or cycle if more numbers than files).

### Single/Direct CLI Message:

```sh
python wa.py msg +919999999999 "Direct message"
python wa.py msg "C:\photo.jpg" +919999999999 "Photo with caption"
```

---

## 🗂️ Excel/CSV Format

* Numbers **column can be any letter (`-col A`) or header name (`-col mobile`)**.
* Always skips header row.
* Accepts numbers with/without country code.

**Sample CSV:**

```csv
mobile
+919111111111
+919222222222
919333333333
```

**Sample Excel:**
Just numbers in column A, or any header you want.

---

## 💡 Tips

* Use **exl** for pure bulk text, **exl\_atg** for bulk attachments+text.
* **All phone numbers are auto-corrected** to international format.
* HTML analytics report is generated after every bulk run.
* **Random delay** between each send for safety.

---

## 🛠️ EXE Build

Turn it into a standalone Windows EXE with:

```sh
pip install pyinstaller
pyinstaller --onefile wa.py
```

---

## 🛡️ Disclaimer

* This project is for **educational/personal use**.
* Use responsibly and respect WhatsApp’s fair usage policy.
* Not affiliated with WhatsApp/Facebook/Meta.

---

## 🙏 Credits

* [Selenium](https://selenium.dev/)
* [webdriver-manager](https://github.com/SergeyPirogov/webdriver_manager)
* [pandas](https://pandas.pydata.org/)
* [openpyxl](https://openpyxl.readthedocs.io/)
* Emoji/UX ideas: [Abhi](mailto:al2025485@gmail.com)

---

## ⭐ Star & Fork if you find it useful!

---

**Happy Automating!**
