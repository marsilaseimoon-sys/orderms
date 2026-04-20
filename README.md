# orderms
Professional Order Management System with Flask, Bootstrap 5, AI Reports &amp; WhatsApp Integration
# 📦 OrderPro — Business Order Management System

<div align="center">

![Python](https://img.shields.io/badge/Python-3.8+-3776AB?style=for-the-badge&logo=python&logoColor=white)
![Flask](https://img.shields.io/badge/Flask-2.x-000000?style=for-the-badge&logo=flask&logoColor=white)
![Bootstrap](https://img.shields.io/badge/Bootstrap-5.3-7952B3?style=for-the-badge&logo=bootstrap&logoColor=white)
![Excel](https://img.shields.io/badge/Excel-Sync-217346?style=for-the-badge&logo=microsoft-excel&logoColor=white)
![License](https://img.shields.io/badge/License-MIT-green?style=for-the-badge)

**A professional, full-featured order management web app built with Flask & Bootstrap 5.**  
Includes AI-powered reports, WhatsApp integration, Excel sync, and a beautiful dark UI.

[✨ Features](#-features) • [🚀 Quick Start](#-quick-start) • [📸 Screenshots](#-screenshots) • [🗂️ Project Structure](#️-project-structure) • [⚙️ Configuration](#️-configuration)

</div>

---

## ✨ Features

| Feature | Description |
|---|---|
| 🔐 **Login System** | Session-based auth with role support (Admin / Manager) |
| 🏠 **Dashboard** | Live stats — total orders, revenue, pending, delivered |
| 📦 **Order Management** | Add, delete, search, filter & update order status |
| 👥 **Customer Register** | Auto-grouped customer profiles with spend history |
| 💬 **WhatsApp Integration** | Send individual or bulk WhatsApp confirmations via `pywhatkit` |
| 📊 **AI Reports** | Charts (revenue trend, donut, top products) + scikit-learn revenue forecast |
| ⏳ **Pending Queue** | One-click bulk WhatsApp send to all pending orders |
| ✅ **Delivered Tracker** | Track completed orders with total revenue |
| 📥 **Excel Import** | Import any `.xlsx` file directly into the app |
| ⚙️ **Settings Page** | Tabbed settings — Profile, Appearance, Notifications, Security, System Info |
| 🌙 **Dark Theme** | Modern dark UI with Bootstrap 5 + custom CSS variables |
| 📱 **Responsive** | Mobile-friendly sidebar with toggle |

---

## 🚀 Quick Start

### 1. Clone the Repository

```bash
git clone https://github.com/YOUR_USERNAME/orderms.git
cd orderms
```

### 2. Install Dependencies

```bash
pip install flask pandas openpyxl pywhatkit scikit-learn matplotlib
```

### 3. Run Setup (First Time Only)

```bash
python setup.py
```

This will:
- ✅ Check all required libraries
- ✅ Create a professional `orders.xlsx` with demo data (3 sheets)
- ✅ Verify all template files exist

### 4. Start the App

```bash
python app.py
```

Browser will open automatically at `http://127.0.0.1:5000` 🎉

---

## 🔐 Login Credentials

| Username | Password | Role |
|---|---|---|
| `admin` | `admin123` | Administrator |
| `manager` | `manager123` | Sales Manager |

> To change credentials, edit the `USERS` dictionary in `app.py`.

---

## 📸 Screenshots

> _Add screenshots of your app here after running it._

| Page | Description |
|---|---|
| Login | Animated dark login with demo credential hints |
| Dashboard | Stats cards + recent orders + quick action cards |
| Orders | Full table with search, filter, status update, WhatsApp send |
| Reports | AI insights + 4 charts + monthly summary table |
| Settings | 6-tab settings page with toggles, profile, security |

---

## 🗂️ Project Structure

```
orderms/
│
├── app.py                  # Main Flask application (all routes + logic)
├── setup.py                # First-time setup — creates orders.xlsx & checks libraries
├── orders.xlsx             # Auto-generated Excel database (3 sheets)
│
└── templates/
    ├── base.html           # Bootstrap 5 sidebar layout (shared by all pages)
    ├── login.html          # Login page with animated background
    ├── index.html          # Dashboard
    ├── orders.html         # All orders with search & filter
    ├── customers.html      # Customer profiles
    ├── reports.html        # AI insights + charts
    ├── pending.html        # Pending orders queue
    ├── whatsapp.html       # WhatsApp send center
    ├── delivered.html      # Delivered orders
    └── settings.html       # 6-tab settings page
```

---

## ⚙️ Configuration

### Change Login Credentials

In `app.py`, find and edit the `USERS` dictionary:

```python
USERS = {
    "admin":   {"password": "your_password", "role": "Admin",   "name": "Your Name"},
    "manager": {"password": "your_password", "role": "Manager", "name": "Manager Name"},
}
```

### WhatsApp Phone Number Format

The app auto-formats Pakistani numbers:

```
03001234567  →  +923001234567
```

For other countries, edit `fmt_phone()` in `app.py`.

### Excel File (orders.xlsx)

The app reads from `orders.xlsx` in the same folder. It auto-detects headers and maps:

| Excel Column | App Field |
|---|---|
| Order ID | OrderID |
| Customer Name | CustomerName |
| Phone | Phone |
| Product / Description | Product |
| Amount (Rs) | Amount |
| Status | Status (Pending / Sent / Delivered) |
| Date | Date |

---

## 📦 Dependencies

```txt
flask
pandas
openpyxl
pywhatkit
scikit-learn
matplotlib
```

Install all at once:

```bash
pip install flask pandas openpyxl pywhatkit scikit-learn matplotlib
```

---

## 🤖 AI Features

The AI reports use **local machine learning** — no API key or internet required:

- **Revenue Forecast** — Linear Regression on monthly revenue data (`scikit-learn`)
- **Top Customer** — Auto-detects highest-spending customer
- **Best Product** — Finds top product by total revenue
- **Delivery Rate** — Calculates % of delivered orders
- **Average Order Value** — Mean order amount

---

## 💬 WhatsApp Integration

Uses `pywhatkit` which opens WhatsApp Web in your browser:

1. Make sure **WhatsApp Web is logged in** before sending
2. Click **💬 Send** on any order — browser opens automatically
3. Use **Send All Pending** for bulk sending

> ⚠️ `pywhatkit` requires a display/browser. It does **not** work on headless servers.

---

## ⚠️ Important Notes

- 📊 All data in `orders.xlsx` is **demo/fictional** — replace with your real data
- 🔒 Change default passwords before using in production
- 💻 Designed to run **locally** on Windows/Mac/Linux
- 🌐 For public deployment, replace `pywhatkit` with WhatsApp Business API

---

## 🛠️ Built With

- **[Flask](https://flask.palletsprojects.com/)** — Python web framework
- **[Bootstrap 5](https://getbootstrap.com/)** — Frontend UI framework
- **[Bootstrap Icons](https://icons.getbootstrap.com/)** — Icon library
- **[pandas](https://pandas.pydata.org/)** — Excel data handling
- **[openpyxl](https://openpyxl.readthedocs.io/)** — Excel file creation & formatting
- **[pywhatkit](https://github.com/Ankit404butfound/PyWhatKit)** — WhatsApp automation
- **[scikit-learn](https://scikit-learn.org/)** — AI/ML forecasting
- **[matplotlib](https://matplotlib.org/)** — Chart generation

---

## 📄 License

This project is licensed under the **MIT License** — feel free to use, modify, and distribute.

---

## 🙌 Contributing

Pull requests are welcome! For major changes, please open an issue first.

1. Fork the repo
2. Create your branch: `git checkout -b feature/amazing-feature`
3. Commit: `git commit -m 'Add amazing feature'`
4. Push: `git push origin feature/amazing-feature`
5. Open a Pull Request

---

<div align="center">
Made with ❤️ using Flask & Bootstrap 5
</div>
